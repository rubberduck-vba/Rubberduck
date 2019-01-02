using System;
using Microsoft.Win32;

namespace Rubberduck.Parsing.ComReflection.TypeLibReflection
{
    // inspired from https://github.com/rossknudsen/Kavod.ComReflection

    public static class RegistryKeyExtensions
    {
        public static string GetKeyName(this RegistryKey key)
        {
            var name = key?.Name;
            return name?.Substring(name.LastIndexOf(@"\", StringComparison.InvariantCultureIgnoreCase) + 1) ?? string.Empty;
        }
    }
}