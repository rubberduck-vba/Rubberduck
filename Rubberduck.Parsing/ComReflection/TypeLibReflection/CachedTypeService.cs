using System;
using System.Collections.Concurrent;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

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
    /// false. Thus, those methods should be wrapped in the <c>TryCachedType</c> methods to ensure that
    /// the repeated invocation will continue to return exactly same <see cref="Type"/>.
    ///
    /// For details on the issue with the <see cref="Type.IsEquivalentTo"/>, refer to:
    /// https://developercommunity.visualstudio.com/content/problem/422208/typeisequivalent-does-not-behave-according-to-the.html
    /// </remarks>
    public interface ICachedTypeService
    {
        bool TryInvalidate(string project, string progId = null);
        bool TryGetCachedType(string progId, out Type type);
        bool TryGetCachedType(string project, string progId, out Type type);
        bool TryGetCachedType(ITypeInfo typeInfo, out Type type);
        bool TryGetCachedType(ITypeInfo typeInfo, string project, out Type type);
        Type TryGetCachedTypeFromEquivalentType(string project, string progId, Type type);
    }

    public class CachedTypeService : ICachedTypeService
    {
        private static readonly ConcurrentDictionary<string, ILibraryTypeCache> TypeCaches;
        private static readonly Lazy<CachedTypeService> LazyInstance;
        private static readonly ITypeLibQueryService QueryService;

        static CachedTypeService()
        {
            TypeCaches = new ConcurrentDictionary<string, ILibraryTypeCache>();
            TypeCaches.TryAdd(string.Empty, new LibraryTypeCache(string.Empty));

            LazyInstance = new Lazy<CachedTypeService>(() => new CachedTypeService());
            QueryService = TypeLibQueryService.Instance;
        }

        /// <summary>
        /// Provided primarily for uses outside the CW's DI, mainly within Rubberduck.Main.
        /// </summary>
        public static ICachedTypeService Instance => LazyInstance.Value;

        public bool TryGetCachedType(string progId, out Type type)
        {
            return TryGetCachedType(string.Empty, progId, out type);
        }

        public bool TryGetCachedType(string project, string progId, out Type type)
        {
            if (TryGetValue(project, progId, out type))
            {
                return type != null;
            }

            type = Type.GetTypeFromProgID(progId);
            if (type == null)
            {
                return type != null;
            }

            if (!TryAddTypeInternal(project, progId, ref type))
            {
                type = null;
            }

            return type != null;
        }

        public bool TryGetCachedType(ITypeInfo typeInfo, out Type type)
        {
            return TryGetCachedType(typeInfo, string.Empty, out type);
        }

        public bool TryGetCachedType(ITypeInfo typeInfo, string project, out Type type)
        {
            var progId = QueryService.GetOrCreateProgIdFromITypeInfo(typeInfo);
            return TryGetCachedType(typeInfo, project, progId, out type);
        }

        private bool TryGetCachedType(ITypeInfo typeInfo, string project, string progId, out Type type)
        {
            if (TryGetValue(project, progId, out type))
            {
                return type != null;
            }

            if (!QueryService.TryGetTypeFromITypeInfo(typeInfo, out type))
            {
                return type != null;
            }

            if (!TryAddTypeInternal(project, progId, ref type))
            {
                return false;
            }

            return type != null;
        }

        public Type TryGetCachedTypeFromEquivalentType(string project, string progId, Type type)
        {
            var cache = TypeCaches.GetOrAdd(project?.ToLowerInvariant() ?? string.Empty, s => new LibraryTypeCache(s));
            return cache.GetOrAdd(progId, type);
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
        private bool TryAddTypeInternal(string project, string progId, ref Type type)
        {
            // Using local function because we don't want to accidentally add types without
            // having went through the logic of checking & obtaining the types.
            bool TryAdd(string progIdToAdd, Type typeToAdd)
            {
                var cache = TypeCaches.GetOrAdd(project?.ToLowerInvariant() ?? string.Empty, s => new LibraryTypeCache(s));
                return cache.AddType(progIdToAdd, typeToAdd);
            }

            // Ensure we do not cache the generic System.__ComObject, which is useless.
            if (type.Name == "__ComObject")
            {
                return QueryService.TryGetTypeInfoFromProgId(progId, out var typeInfo) 
                       && TryGetCachedType(typeInfo, project?.ToLowerInvariant() ?? string.Empty, progId, out type);
            }

            if (!TryAdd(progId, type))
            {
                return false;
            }

            return type.GetInterfaces()
                .Where(face => face.FullName != null)
                .All(face => TryAdd(face.FullName, face));
        }

        private static bool TryGetValue(string project, string progId, out Type type)
        {
            if (TypeCaches.TryGetValue(project?.ToLowerInvariant() ?? string.Empty, out var cache))
            {
                return cache.TryGetType(progId, out type);
            }

            type = null;
            return false;
        }

        public bool TryInvalidate(string project, string progId = null)
        {
            if (TypeCaches.TryGetValue(project?.ToLowerInvariant() ?? string.Empty, out var cache))
            {
                if (!string.IsNullOrWhiteSpace(progId))
                {
                    return cache.Remove(progId);
                }
                else
                {
                    return TypeCaches.TryRemove(cache.Key, out _);
                }
            }

            return false;
        }
    }
}