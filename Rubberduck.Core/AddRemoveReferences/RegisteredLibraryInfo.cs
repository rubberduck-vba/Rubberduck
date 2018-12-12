using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.AddRemoveReferences
{
    public struct RegisteredLibraryKey
    {
        public Guid Guid { get; }
        public int Major { get; }
        public int Minor { get; }

        public RegisteredLibraryKey(Guid guid, int major, int minor)
        {
            Guid = guid;
            Major = major;
            Minor = minor;
        }
    }

    public class RegisteredLibraryInfo
    {
        private static readonly Dictionary<int, string> NativeLocaleNames = new Dictionary<int, string>
        {
            { 0, Resources.RubberduckUI.References_DefaultLocale }
        };

        public RegisteredLibraryKey UniqueId { get; }
        public string Name { get; set; }
        public Guid Guid { get; set; }
        public string Description { get; set; }
        public string Version => $"{Major}.{Minor}";
        public string FullPath => string.IsNullOrEmpty(FullPath32) || Has64BitVersion && Environment.Is64BitProcess ? FullPath64 : FullPath32;
        public int Major { get; set; }
        public int Minor { get; set; }
        public int LocaleId { get; set; }

        public string LocaleName
        {
            get
            {
                if (NativeLocaleNames.ContainsKey(LocaleId))
                {
                    return NativeLocaleNames[LocaleId];
                }

                try
                {
                    var name = CultureInfo.GetCultureInfo(LocaleId).NativeName;
                    NativeLocaleNames.Add(LocaleId, name);
                    return name;
                }
                catch
                {
                    NativeLocaleNames.Add(LocaleId, Resources.RubberduckUI.References_DefaultLocale);
                    return Resources.RubberduckUI.References_DefaultLocale;
                }
            }
        }

        public LIBFLAGS Flags { get; set; }

        private string FullPath32 { get; }
        private string FullPath64 { get; }
        public bool Has32BitVersion => !string.IsNullOrEmpty(FullPath32);
        public bool Has64BitVersion => !string.IsNullOrEmpty(FullPath64);

        public int? Priority => null;

        public RegisteredLibraryInfo(Guid guid, string description, string version, string path32, string path64)
        {
            Guid = guid;

            var majorMinor = version.Split('.');
            if (majorMinor.Length == 2)
            {
                Major = int.TryParse(majorMinor[0], NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out var major) ? major : 0;
                Minor = int.TryParse(majorMinor[1], NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out var minor) ? minor : 0;
            }

            FullPath32 = path32 ?? string.Empty;
            FullPath64 = path64 ?? string.Empty;

            Description = !string.IsNullOrEmpty(description) ? description : Path.GetFileNameWithoutExtension(FullPath);
            UniqueId = new RegisteredLibraryKey(guid, Major, Minor);
        }

        public override string ToString() => Description;
    }
}