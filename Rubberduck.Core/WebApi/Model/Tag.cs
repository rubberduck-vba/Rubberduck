using System;
using System.Collections.Generic;

namespace Rubberduck.Core.WebApi.Model
{
    public class Tag : Entity
    {
        public string Name { get; set; }
        public Version Version => new Version(Name.Substring(1));
        public DateTime DateCreated { get; set; }
        public string InstallerDownloadUrl { get; set; }
        public int InstallerDownloads { get; set; }
        public bool IsPreRelease { get; set; }

        public virtual ICollection<TagAsset> TagAssets { get; set; } = new List<TagAsset>();
    }
}
