using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Win32;

namespace Rubberduck.Parsing.ComReflection.TypeLibReflection
{
    public class TypeLibQueryService
    {
        [DllImport("ole32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, PreserveSig = true)]
        private static extern int CLSIDFromProgID(string lpszProgID, out Guid lpclsid);

        [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, PreserveSig = true)]
        private static extern int LoadTypeLib(string fileName, out ITypeLib typeLib);

        private static readonly Lazy<TypeLibQueryService> LazyInstance = new Lazy<TypeLibQueryService>();
        private static readonly RegisteredLibraryFinderService Finder = new RegisteredLibraryFinderService();

        /// <summary>
        /// Provided primarily for uses outside the CW's DI, mainly within Rubberduck.Main.
        /// </summary>
        public static TypeLibQueryService Instance => LazyInstance.Value;

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
            return true;
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

            var file = Path.GetFileName(path);
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
