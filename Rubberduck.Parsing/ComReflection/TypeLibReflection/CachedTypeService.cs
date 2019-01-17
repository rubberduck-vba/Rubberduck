using System;
using System.Collections.Concurrent;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor.Utility;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;

namespace Rubberduck.Parsing.ComReflection.TypeLibReflection
{
    /// <summary>
    /// Provide caching service for types that should be considered equivalent.
    /// </summary>
    /// <remarks>
    /// The provider aims to work around a deficiency in the <see cref="Type.IsEquivalentTo"/>, particularly for
    /// COM interop types. The issue is that when we create a <see cref="Type"/> derived from methods such as
    /// <see cref="Marshal.GetTypeForITypeInfo"/> or <see cref="Type.GetTypeFromProgID(string)"/>, new types are
    /// returned for each invocation, even for the same ProgID or ITypeInfo. That will cause problems later such
    /// as being unable to cast an instance from one type to another, even though they are based on exactly the
    /// same ProgID/ITypeInfo/etc.. In those cases, the <see cref="Type.IsEquivalentTo"/> incorrectly returns
    /// false. Thus, those methods should be wrapped in the <see cref="TryCacheType"/> methods to ensure that
    /// the repeated invocation will continue to return exactly same <see cref="Type"/>.
    ///
    /// For details on the issue with the <see cref="Type.IsEquivalentTo"/>, refer to:
    /// https://developercommunity.visualstudio.com/content/problem/422208/typeisequivalent-does-not-behave-according-to-the.html
    /// </remarks>
    public interface ICachedTypeService
    {
        bool TryGetCachedType(string progId, out Type type);
        bool TryGetCachedType(string progId, string project, out Type type);
        bool TryGetCachedType(ITypeInfo typeInfo, out Type type);
        bool TryGetCachedType(ITypeInfo typeInfo, string project, out Type type);
    }

    public class CachedTypeService : ICachedTypeService
    {
        private static readonly ConcurrentDictionary<string, Type> TypeCache;
        private static readonly Lazy<CachedTypeService> LazyInstance;
        private static readonly ITypeLibQueryService QueryService;

        static CachedTypeService()
        {
            TypeCache = new ConcurrentDictionary<string, Type>();
            LazyInstance = new Lazy<CachedTypeService>(() => new CachedTypeService());
            QueryService = TypeLibQueryService.Instance;
        }

        /// <summary>
        /// Provided primarily for uses outside the CW's DI, mainly within Rubberduck.Main.
        /// </summary>
        public static ICachedTypeService Instance => LazyInstance.Value;

        public bool TryGetCachedType(string progId, out Type type)
        {
            return TryGetCachedType(progId, null, out type);
        }

        public bool TryGetCachedType(string progId, string project, out Type type)
        {
            var key = CreateQualifiedIdentifier(progId, project);
            if (!TypeCache.TryGetValue(key, out type))
            {
                type = Type.GetTypeFromProgID(progId);
                if (type != null)
                {
                    if (!TryAddTypeInternal(progId, project, ref type))
                    {
                        type = null;
                    }
                }
            }

            return type != null;
        }

        public bool TryGetCachedType(ITypeInfo typeInfo, out Type type)
        {
            return TryGetCachedType(typeInfo, null, out type);
        }

        public bool TryGetCachedType(ITypeInfo typeInfo, string project, out Type type)
        {
            typeInfo.GetTypeAttr(out var pAttr);
            if (pAttr != IntPtr.Zero)
            {
                using (DisposalActionContainer.Create(pAttr, typeInfo.ReleaseTypeAttr))
                {
                    var attr = Marshal.PtrToStructure<TYPEATTR>(pAttr);
                    var clsid = attr.guid;
                    if (QueryService.TryGetProgIdFromClsid(clsid, out var progId))
                    {
                        return TryGetCachedType(typeInfo, progId, project, out type);
                    }
                }
            }

            var typeName = Marshal.GetTypeInfoName(typeInfo);
            typeInfo.GetContainingTypeLib(out var typeLib, out _);
            var libName = Marshal.GetTypeLibName(typeLib);

            return TryGetCachedType(typeInfo, string.Concat(libName, ".", typeName), project, out type);
        }

        private bool TryGetCachedType(ITypeInfo typeInfo, string progId, string project, out Type type)
        {
            var key = CreateQualifiedIdentifier(progId, project);
            if (TypeCache.TryGetValue(key, out type))
            {
                return type != null;
            }

            var ptr = Marshal.GetComInterfaceForObject(typeInfo, typeof(ITypeInfo));
            if (ptr == IntPtr.Zero)
            {
                return false;
            }

            using (DisposalActionContainer.Create(ptr, x => Marshal.Release(x)))
            {
                type = Marshal.GetTypeForITypeInfo(ptr);
                if (type == null)
                {
                    return false;
                }

                if (!TryAddTypeInternal(progId, project, ref type))
                {
                    return false;
                }
            }

            return type != null;
        }

        /// <summary>
        /// Because a <see cref="Type"/> can have several interfaces and those may be further used in
        /// downstream operations, it's important to also cache those interfaces to ensure we do not
        /// return a different type for a given interface that's implemented by the cached type.
        ///
        /// Additionally, we ensure that we do not cache any <see cref="System.__ComObject"/> types
        /// as those are not useful in production. In that case, we must discover the type library
        /// using the <see cref="TypeLibQueryService"/> and call <see cref="Marshal.GetTypeForITypeInfo"/>.
        /// </summary>
        /// <returns>True if the type and all its interface were added. False otherwise</returns>
        private bool TryAddTypeInternal(string progId, string project, ref Type type)
        {
            // Ensure we do not cache the generic System.__ComObject, which is useless.
            if (type.Name == "__ComObject")
            {
                return QueryService.TryGetTypeInfoFromProgId(progId, out var typeInfo) 
                       && TryGetCachedType(typeInfo, progId, project, out type);
            }

            if (!TypeCache.TryAdd(CreateQualifiedIdentifier(progId, project), type))
            {
                return false;
            }

            return type.GetInterfaces()
                .Where(face => face.FullName != null)
                .All(face => TypeCache.TryAdd(CreateQualifiedIdentifier(face.FullName, project), face));
        }

        /// <summary>
        /// Creates a qualified identifier to uniquely identify a cached type, with optional scoping. Case insensitive.
        /// </summary>
        /// <remarks>
        /// A typical use is to distinguish the types by its ProgID / <see cref="Type.FullName"/>. However,
        /// if a type comes from a private project there is a potential for a collision. In that case, the
        /// optional project should be filled in.
        /// </remarks>
        /// <param name="progId">Unique name for the type.</param>
        /// <param name="project">Indicates whether the type belongs to a privately scoped project. Leave null to indicate it's global</param>
        /// <returns>A fully qualified identifier</returns>
        private static string CreateQualifiedIdentifier(string progId, string project)
        {
            return string.Concat(project?.ToLowerInvariant(), "::", progId.ToLowerInvariant());
        }
    }
}