using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Win32;

namespace Rubberduck.AddRemoveReferences
{
    public interface IRegisteredLibraryFinderService
    {
        IEnumerable<RegisteredLibraryInfo> FindRegisteredLibraries();
    }

    // inspired from https://github.com/rossknudsen/Kavod.ComReflection
    public class RegisteredLibraryFinderService : IRegisteredLibraryFinderService
    {
        private static readonly List<string> IgnoredKeys = new List<string> { "FLAGS", "HELPDIR" };

        public IEnumerable<RegisteredLibraryInfo> FindRegisteredLibraries()
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

                        foreach (var lcid in versionKey.GetSubKeyNames().Where(key => !IgnoredKeys.Contains(key)))
                        {
                            if (!int.TryParse(lcid, out var id))
                            {
                                continue;
                            }
                            using (var paths = versionKey.OpenSubKey(lcid))
                            {
                                string bit32;
                                string bit64;
                                using (var win32 = paths?.OpenSubKey("win32"))
                                {
                                    bit32 = win32?.GetValue(string.Empty)?.ToString() ?? string.Empty;
                                }
                                using (var win64 = paths?.OpenSubKey("win64"))
                                {
                                    bit64 = win64?.GetValue(string.Empty)?.ToString() ?? string.Empty;
                                }

                                yield return new RegisteredLibraryInfo(guid, name, version, bit32, bit64)
                                {
                                    Flags = flagValue,
                                    LocaleId = id
                                };
                            }
                        }
                    }
                }
            }
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

    public static class RegistryKeyExtensions
    {
        public static string GetKeyName(this RegistryKey key)
        {
            var name = key?.Name;
            return name?.Substring(name.LastIndexOf(@"\", StringComparison.InvariantCultureIgnoreCase) + 1) ?? string.Empty;
        }
    }
}