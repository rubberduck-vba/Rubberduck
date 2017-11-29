using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for ignoring specific inspection results from a specified set of inspections.
    /// </summary>
    public sealed class IgnoreAnnotation : AnnotationBase
    {
        public IgnoreAnnotation(
            QualifiedSelection qualifiedSelection,
            IEnumerable<string> parameters)
            : base(AnnotationType.Ignore, qualifiedSelection)
        {
            InspectionNames = parameters;
        }

        public IEnumerable<string> InspectionNames { get; }

        public bool IsIgnored(string inspectionName)
        {
            return InspectionNames.Contains(inspectionName);
        }

        public override string ToString()
        {
            return $"Ignored inspections: {string.Join(", ", InspectionNames)}";
        }
    }
}
