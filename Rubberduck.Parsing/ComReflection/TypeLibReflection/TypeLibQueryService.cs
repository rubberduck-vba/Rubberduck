using System;
using System.IO.Abstractions;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Win32;
using Rubberduck.InternalApi.Common;
using Rubberduck.VBEditor.Utility;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;

namespace Rubberduck.Parsing.ComReflection.TypeLibReflection
{
    public interface ITypeLibQueryService
    {
        bool TryGetTypeFromITypeInfo(ITypeInfo typeInfo, out Type type);
        bool TryGetProgIdFromClsid(Guid clsid, out string progId);
        bool TryGetTypeInfoFromProgId(string progId, out ITypeInfo typeInfo);
        string GetOrCreateProgIdFromITypeInfo(ITypeInfo typeInfo);
    }

    public class TypeLibQueryService : ITypeLibQueryService
    {
        [DllImport("ole32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, PreserveSig = true)]
        private static extern int CLSIDFromProgID(string lpszProgID, out Guid lpclsid);

        [DllImport("ole32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, PreserveSig = true)]
        private static extern int ProgIDFromCLSID([In]ref Guid clsid, [MarshalAs(UnmanagedType.LPWStr)]out string lplpszProgID);

        [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, PreserveSig = true)]
        private static extern int LoadTypeLib(string fileName, out ITypeLib typeLib);

        private static readonly Lazy<TypeLibQueryService> LazyInstance = new Lazy<TypeLibQueryService>();
        private static readonly RegisteredLibraryFinderService Finder = new RegisteredLibraryFinderService();

        /// <summary>
        /// Provided primarily for uses outside the CW's DI, mainly within Rubberduck.Main.
        /// </summary>
        public static ITypeLibQueryService Instance => LazyInstance.Value;

        public bool TryGetTypeFromITypeInfo(ITypeInfo typeInfo, out Type type)
        {
            type = null;
            var ptr = Marshal.GetComInterfaceForObject(typeInfo, typeof(ITypeInfo));
            if (ptr == IntPtr.Zero)
            {
                return false;
            }

            using (DisposalActionContainer.Create(ptr, x => Marshal.Release(x)))
            {
                type = Marshal.GetTypeForITypeInfo(ptr);
            }

            return type != null;
        }

        public string GetOrCreateProgIdFromITypeInfo(ITypeInfo typeInfo)
        {
            typeInfo.GetTypeAttr(out var pAttr);
            if (pAttr != IntPtr.Zero)
            {
                using (DisposalActionContainer.Create(pAttr, typeInfo.ReleaseTypeAttr))
                {
                    var attr = Marshal.PtrToStructure<TYPEATTR>(pAttr);
                    var clsid = attr.guid;
                    if (TryGetProgIdFromClsid(clsid, out var progId))
                    {
                        return progId;
                    }
                }
            }

            var typeName = Marshal.GetTypeInfoName(typeInfo);
            typeInfo.GetContainingTypeLib(out var typeLib, out _);
            var libName = Marshal.GetTypeLibName(typeLib);

            return string.Concat(libName, ".", typeName);
        }

        public bool TryGetProgIdFromClsid(Guid clsid, out string progId)
        {
            return ProgIDFromCLSID(ref clsid, out progId) == 0;
        }

        public bool TryGetTypeInfoFromProgId(string progId, out ITypeInfo typeInfo)
        {
            typeInfo = null;
            if (CLSIDFromProgID(progId, out var clsid) != 0)
            {
                return false;
            }

            if (!TryGetTypeLibFromClsid(clsid, out var lib))
            {
                return false;
            }
            
            lib.GetTypeInfoOfGuid(ref clsid, out typeInfo);
            return typeInfo != null;
        }

        private static bool TryGetTypeLibFromClsid(Guid clsid, out ITypeLib lib)
        {
            lib = null;

            using (var clsidKey = Registry.ClassesRoot.OpenSubKey($"CLSID\\{clsid:B}"))
            {
                if (clsidKey == null)
                {
                    return false;
                }

                if (!TryLoadTypeLibFromPath(TryGetTypeLibPath, clsidKey, out lib))
                {
                    return true;
                }

                if (TryLoadTypeLibFromPath(TryGetInProcServerPath, clsidKey, out lib))
                {
                    return true;
                }

                if (TryLoadTypeLibFromPath(TryGetLocalServerPath, clsidKey, out lib))
                {
                    return true;
                }

                return false;
            }
        }

        private delegate bool GetPathFunction(RegistryKey clsidKey, out string path);
        private static bool TryLoadTypeLibFromPath(GetPathFunction getPathFunction, RegistryKey clsidKey, out ITypeLib lib)
        {
            lib = null;
            if (!getPathFunction(clsidKey, out var path))
            {
                return false;
            }

            if (LoadTypeLib(path, out lib) == 0)
            {
                return true;
            }

            var file = FileSystemProvider.FileSystem.Path.GetFileName(path);
            return LoadTypeLib(file, out lib) == 0;
        }

        private static bool TryGetTypeLibPath(RegistryKey clsidKey, out string path)
        {
            path = null;

            using (var clsidTypeLibKey = clsidKey.OpenSubKey("TypeLib"))
            {
                if (clsidTypeLibKey == null)
                {
                    return false;
                }

                if (Guid.TryParseExact(((string) clsidTypeLibKey.GetValue(null)).ToLowerInvariant(), "B", out var libGuid) 
                    && Finder.TryGetRegisteredLibraryInfo(libGuid, out var info))
                {
                    path = info.FullPath;
                }
            }

            return !string.IsNullOrWhiteSpace(path);
        }

        private static bool TryGetInProcServerPath(RegistryKey clsidKey, out string path)
        {
            using (var procServerKey = clsidKey.OpenSubKey("InprocServer32"))
            {
                if (procServerKey != null)
                {
                    path = procServerKey.GetValue(null) as string;
                    return true;
                }

                path = null;
                return false;
            }
        }

        private static bool TryGetLocalServerPath(RegistryKey clsidKey, out string path)
        {
            using (var localServerKey = clsidKey.OpenSubKey("LocalServer32"))
            {
                if (localServerKey != null)
                {
                    path = localServerKey.GetValue(null) as string;
                    return true;
                }

                path = null;
                return false;
            }
        }
    }
}
