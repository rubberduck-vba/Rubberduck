using System;
using System.Collections.Generic;
using Microsoft.Win32;

namespace Rubberduck.AddRemoveReferences
{
    public interface IRegisteredLibraryFinderService
    {
        IEnumerable<ReferenceModel> FindRegisteredLibraries();
    }

    public class RegisteredLibraryFinderService : IRegisteredLibraryFinderService
    {
        private readonly bool _use64BitPaths;

        // inspired from https://github.com/rossknudsen/Kavod.ComReflection
        public RegisteredLibraryFinderService(bool use64BitPaths)
        {
            _use64BitPaths = use64BitPaths;
        }

        public IEnumerable<ReferenceModel> FindRegisteredLibraries()
        {
            using (var typelibSubKey = Registry.ClassesRoot.OpenSubKey("TypeLib"))
            {
                if (typelibSubKey == null) { yield break; }

                var typeLibs = EnumerateSubKeys(typelibSubKey);
                foreach (var key in typeLibs)
                {
                    var info = new RegisteredLibraryInfo
                    {
                        Guid = typelibSubKey.Name,
                        Name = key.GetValue(string.Empty).ToString(),
                        FullPath = GetLibraryPath(key),
                        Version = new Version(key.Name),
                    };
                }
            }
        }

        private IEnumerable<RegistryKey> EnumerateSubKeys(RegistryKey key)
        {
            foreach (var keyName in key.GetSubKeyNames())
            {
                using (var subKey = key.OpenSubKey(keyName))
                {
                    yield return subKey;
                }
            }
        }

        private string GetLibraryPath(RegistryKey versionSubKey)
        {
            using (var subKey = versionSubKey.OpenSubKey("0"))
            {
                if (subKey == null) { return string.Empty; }
                if (_use64BitPaths)
                {
                    // if host process is 64-bit, use x64 paths when possible
                    var win64 = subKey.OpenSubKey("win64");
                    if (win64 != null)
                    {
                        return win64.GetValue(string.Empty).ToString();
                    }
                }

                var win32 = subKey.OpenSubKey("win32");
                if (win32 != null)
                {
                    return win32.GetValue(string.Empty).ToString();
                }

                return string.Empty;
            }
        }
    }
}