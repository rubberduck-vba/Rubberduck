using System.Collections.Generic;

namespace Rubberduck.Core.WebApi.Model
{
    public class FeatureItem : Entity
    {
        public int FeatureId { get; set; }

        public string Name { get; set; }

        public string Title { get; set; }
        public string Description { get; set; }
        public bool IsNew { get; set; }
        public bool IsDiscontinued { get; set; }
        public bool IsHidden { get; set; }
        public int? TagAssetId { get; set; }
        public string XmlDocSourceObject { get; set; }
        public string XmlDocTabName { get; set; }
        public string XmlDocMetadata { get; set; }
        public string XmlDocSummary { get; set; }
        public string XmlDocInfo { get; set; }
        public string XmlDocRemarks { get; set; }

        public virtual Feature Feature { get; set; }
        public virtual TagAsset TagAsset { get; set; }
        public virtual ICollection<Example> Examples { get; set; } = new List<Example>();
    }
}
