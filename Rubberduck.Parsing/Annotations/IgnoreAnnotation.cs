using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for ignoring specific inspection results from a specified set of inspections.
    /// </summary>
    public sealed class IgnoreAnnotation : AnnotationBase
    {
        private readonly IEnumerable<string> _inspectionNames;

        public IgnoreAnnotation(
            QualifiedSelection qualifiedSelection,
            VBAParser.AnnotationContext context,
            IEnumerable<string> parameters)
            : base(AnnotationType.Ignore, qualifiedSelection, context)
        {
            _inspectionNames = parameters;
        }

        public IEnumerable<string> InspectionNames => _inspectionNames;

        public bool IsIgnored(string inspectionName)
        {
            return _inspectionNames.Contains(inspectionName);
        }

        public override bool AllowMultiple { get; } = true;

        public override string ToString()
        {
            return $"Ignored inspections: {string.Join(", ", _inspectionNames)}";
        }
    }
}
