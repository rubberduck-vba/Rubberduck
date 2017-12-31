using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Microsoft.Win32;

namespace Rubberduck.AddRemoveReferences
{
    public static class RegistryKeyExtensions
    {
        public static string GetKeyName(this RegistryKey key)
        {
            var name = key?.Name;
            return name?.Substring(name.LastIndexOf(@"\", StringComparison.InvariantCultureIgnoreCase) + 1) ?? string.Empty;
        }
    }

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
                foreach (var guidKey in typeLibs)
                {
                    foreach (var versionKey in EnumerateSubKeys(guidKey))
                    {
                        var name = versionKey.GetValue(string.Empty)?.ToString();
                        var path = GetLibraryPath(versionKey, out int subKey);
                        if (name == null || string.IsNullOrEmpty(path))
                        {
                            continue;
                        }

                        var flagValue = 0;
                        using (var flagsKey = versionKey.OpenSubKey("FLAGS"))
                        {
                            if (flagsKey != null)
                            {
                                var flags = flagsKey.GetValue(string.Empty)?.ToString() ?? "0";
                                int.TryParse(flags, out flagValue);
                            }
                        }

                        var info = new RegisteredLibraryInfo
                        {
                            Guid = GetGuidFromKeyName(guidKey.GetKeyName()),
                            Name = name,
                            FullPath = path,
                            Version = GetVersionFromKeyName(versionKey.GetKeyName()),
                            Flags = flagValue, // not sure if useful for further filtering
                            SubKey = subKey,   // not sure if useful for further filtering
                        };
                        yield return new ReferenceModel(info);
                    }
                }
            }
        }

        private string GetGuidFromKeyName(string name)
        {
            if (Guid.TryParseExact(name.ToLowerInvariant(), "B", out Guid guid))
            {
                return guid.ToString();
            }

            return string.Empty;
        }

        private Version GetVersionFromKeyName(string name)
        {
            if (Version.TryParse(name, out Version version))
            {
                return version;
            }

            return new Version(0, 0);
        }

        private static readonly HashSet<string> Extensions = new HashSet<string> { "tlb", "olb", "dll" };

        private string GetLibraryPath(RegistryKey versionSubKey, out int subKeyValue)
        {
            var result = string.Empty;
            foreach (var subKeyName in versionSubKey.GetSubKeyNames())
            {
                if (!int.TryParse(subKeyName, out subKeyValue))
                {
                    continue;
                }

                using (var subKey = versionSubKey.OpenSubKey(subKeyName))
                {
                    if (subKey == null)
                    {
                        continue;
                    }

                    var path = string.Empty;
                    if (_use64BitPaths)
                    {
                        // if host process is 64-bit, use x64 paths when possible
                        var win64 = subKey.OpenSubKey("win64");
                        if (win64 != null)
                        {
                            path = win64.GetValue(string.Empty).ToString();
                        }
                    }

                    if (string.IsNullOrEmpty(path))
                    {
                        var win32 = subKey.OpenSubKey("win32");
                        if (win32 != null)
                        {
                            path = win32.GetValue(string.Empty).ToString();
                        }
                    }

                    if (string.IsNullOrEmpty(path))
                    {
                        continue;
                    }

                    var file = new FileInfo(path);
                    var fullPath = file.FullName;
                    if (Extensions.Contains(GetLibraryExtension(fullPath)))
                    {
                        result = fullPath;
                        break;
                    }
                }
            }
            return result;
        }

        private string GetLibraryExtension(string fullPath)
        {
            var lastBackslashIndex = fullPath.LastIndexOf(@"\", StringComparison.OrdinalIgnoreCase);
            var lastDotIndex = fullPath.LastIndexOf(".", StringComparison.OrdinalIgnoreCase);
            var result = lastBackslashIndex > lastDotIndex 
                ? fullPath.Substring(lastDotIndex + 1, lastBackslashIndex - lastDotIndex) 
                : fullPath.Substring(lastDotIndex + 1);
            return result;
        }

        private IEnumerable<RegistryKey> EnumerateSubKeys(RegistryKey key)
        {
            foreach (var keyName in key.GetSubKeyNames())
            {
                using (var subKey = key.OpenSubKey(keyName))
                {
                    if (subKey != null)
                    {
                        yield return subKey;
                    }
                }
            }
        }
    }
}