using System;
using System.Collections.Generic;

namespace Rubberduck.VersionCheck
{
    public class Tag
    {
        public string Name { get; set; }
        public DateTime DateCreated { get; set; }
        public string InstallerDownloadUrl { get; set; }
        public int InstallerDownloads { get; set; }
        public bool IsPreRelease { get; set; }

        public Version Version => new Version(IsPreRelease
            ? Name.Substring("prerelease-v".Length)
            : Name.Substring("v".Length));

        public virtual ICollection<TagAsset> TagAssets { get; set; } = new List<TagAsset>();
    }
}
