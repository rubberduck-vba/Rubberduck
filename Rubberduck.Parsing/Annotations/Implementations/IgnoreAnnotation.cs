using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for ignoring specific inspection results from a specified set of inspections.
    /// </summary>
    [Annotation("Ignore", AnnotationTarget.General, true)]
    public sealed class IgnoreAnnotation : AnnotationBase
    {
        public IgnoreAnnotation(
            QualifiedSelection qualifiedSelection,
            VBAParser.AnnotationContext context,
            IEnumerable<string> parameters)
            : base(qualifiedSelection, context)
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
