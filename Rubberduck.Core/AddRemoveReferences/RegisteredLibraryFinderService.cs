using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices.ComTypes;
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

        private static readonly List<string> IgnoredKeys = new List<string> { "FLAGS", "HELPDIR" };

        public IEnumerable<ReferenceModel> FindRegisteredLibraries()
        {
            using (var typelibSubKey = Registry.ClassesRoot.OpenSubKey("TypeLib"))
            {
                if (typelibSubKey == null) { yield break; }

                foreach (var guidKey in EnumerateSubKeys(typelibSubKey))
                {
                    var guid = Guid.TryParseExact(guidKey.GetKeyName().ToLowerInvariant(), "B", out var clsid)
                        ? clsid
                        : Guid.Empty;
                    
                    foreach (var versionKey in EnumerateSubKeys(guidKey))
                    {
                        var name = versionKey.GetValue(string.Empty)?.ToString();
                        var version = versionKey.GetKeyName();

                        var flagValue = (LIBFLAGS)0;
                        using (var flagsKey = versionKey.OpenSubKey("FLAGS"))
                        {
                            if (flagsKey != null)
                            {
                                var flags = flagsKey.GetValue(string.Empty)?.ToString() ?? "0";
                                Enum.TryParse(flags, out flagValue);
                            }
                        }

                        foreach (var lcid in EnumerateSubKeys(versionKey))
                        {
                            if (IgnoredKeys.Contains(lcid.GetValue(string.Empty)?.ToString()))
                            {
                                continue;
                            }

                            string bit32;
                            string bit64;
                            using (var win32 = lcid.OpenSubKey("win32"))
                            {
                                bit32 = win32?.GetValue(string.Empty)?.ToString() ?? string.Empty;
                            }
                            using (var win64 = lcid.OpenSubKey("win64"))
                            {
                                bit64 = win64?.GetValue(string.Empty)?.ToString() ?? string.Empty;
                            }
                            var info = new RegisteredLibraryInfo(guid, name, version, bit32, bit64)
                            {
                                Flags = flagValue,
                            };

                            yield return new ReferenceModel(info);
                        }
                    }
                }
            }
        }

        //private static readonly HashSet<string> Extensions = new HashSet<string> { "tlb", "olb", "dll" };

        //private string GetLibraryExtension(string fullPath)
        //{
        //    var lastBackslashIndex = fullPath.LastIndexOf(@"\", StringComparison.OrdinalIgnoreCase);
        //    var lastDotIndex = fullPath.LastIndexOf(".", StringComparison.OrdinalIgnoreCase);
        //    var result = lastBackslashIndex > lastDotIndex 
        //        ? fullPath.Substring(lastDotIndex + 1, lastBackslashIndex - lastDotIndex) 
        //        : fullPath.Substring(lastDotIndex + 1);
        //    return result;
        //}

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