using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Core.WebApi.Model
{
    public class Feature : Entity
    {
        public static readonly string[] ProtectedFeatures = new[]
        {
            "CodeInspections", // top-level
            "Inspections", // sub-feature with xmldoc items
            "QuickFixes", // sub-feature with xmldoc items
            "Annotations" // top-level with xmldoc items
        };

        public bool IsProtected => ProtectedFeatures.Contains(Name);

        /// <summary>
        /// Refers to the <c>Id</c> of the parent feature if applicable; <c>null</c> otherwise.
        /// </summary>
        public int? ParentId { get; set; }
        /// <summary>
        /// A short, unique identifier string.
        /// </summary>
        /// <remarks>
        /// While values are PascalCased, this value should be URL-encoded regardless, if used as part of a valid URI.
        /// </remarks>
        public string Name { get; set; }
        /// <summary>
        /// The name of the feature.
        /// </summary>
        public string Title { get; set; }
        /// <summary>
        /// A short paragraph (2-3 sentences) that describes the feature.
        /// </summary>
        public string ElevatorPitch { get; set; }
        /// <summary>
        /// Markdown content for a detailed description of the feature.
        /// </summary>
        public string Description { get; set; }
        /// <summary>
        /// Indicates whether this feature exists in the [next] branch but not yet in the [main] one.
        /// </summary>
        public bool IsNew { get; set; }
        /// <summary>
        /// Indicates whether this feature should be shown.
        /// </summary>
        public bool IsHidden { get; set; }
        /// <summary>
        /// An integer value that is used for sorting the features.
        /// </summary>
        /// <remarks>
        /// Client should use the <c>Name</c> as a 2nd sort level.
        /// </remarks>
        public int SortOrder { get; set; }
        /// <summary>
        /// The name of the .xml file this record comes from.
        /// </summary>
        public string XmlDocSource { get; set; }

        public virtual Feature ParentFeature { get; set; }
        public virtual ICollection<Feature> SubFeatures { get; set; } = new List<Feature>();
        public virtual ICollection<FeatureItem> FeatureItems { get; set; } = new List<FeatureItem>();

        public override bool Equals(object obj)
        {
            if (obj is null)
            {
                return false;
            }

            var other = obj as Feature;
            {
                return other?.Name == Name;
            }
        }
        public override int GetHashCode() => HashCode.Compute(Name);
    }
}
