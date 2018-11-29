using System;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.AddRemoveReferences
{
    public class RegisteredLibraryInfo : IReferenceInfo
    {
        public string Name { get; set; }
        public Guid Guid { get; set; }
        public string Description { get; set; }
        public string Version => $"{Major}.{Minor}";
        public string FullPath => string.IsNullOrEmpty(FullPath32) || Has64BitVersion && Environment.Is64BitProcess ? FullPath64 : FullPath32;
        public int Major { get; set; }
        public int Minor { get; set; }
        public LIBFLAGS Flags { get; set; }

        private string FullPath32 { get; }
        private string FullPath64 { get; }
        public bool Has32BitVersion => !string.IsNullOrEmpty(FullPath32);
        public bool Has64BitVersion => !string.IsNullOrEmpty(FullPath64);

        public RegisteredLibraryInfo(Guid guid, string name, string version, string path32, string path64)
        {
            Guid = guid;

            var majorMinor = version.Split('.');
            if (majorMinor.Length == 2)
            {
                Major = int.TryParse(majorMinor[0], NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out var major) ? major : 0;
                Minor = int.TryParse(majorMinor[1], NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out var minor) ? minor : 0;
            }

            FullPath32 = path32;
            FullPath64 = path64;

            Name = !string.IsNullOrEmpty(name) ? name : Path.GetFileNameWithoutExtension(FullPath);
        }

        public override string ToString() => Name;
    }
}