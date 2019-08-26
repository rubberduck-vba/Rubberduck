using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// This annotation allows ignoring inspection results of defined inspections for a whole module
    /// </summary>
    [Annotation("IgnoreModule", AnnotationTarget.Module, true)]
    public sealed class IgnoreModuleAnnotation : AnnotationBase
    {
        public IgnoreModuleAnnotation(QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IEnumerable<string> parameters)
            : base(qualifiedSelection, context)
        {
            InspectionNames = parameters;
        }

        public IEnumerable<string> InspectionNames { get; }

        public bool IsIgnored(string inspectionName)
        {
            return !InspectionNames.Any() || InspectionNames.Contains(inspectionName);
        }

        public override string ToString()
        {
            return $"Ignored inspections: {string.Join(", ", InspectionNames)}";
        }
    }
}