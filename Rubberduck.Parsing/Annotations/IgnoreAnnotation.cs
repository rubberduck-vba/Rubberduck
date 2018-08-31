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
        private readonly IEnumerable<string> _inspectionNames;

        public IgnoreAnnotation(
            QualifiedSelection qualifiedSelection,
            IEnumerable<string> parameters)
            : base(AnnotationType.Ignore, qualifiedSelection)
        {
            _inspectionNames = parameters;
        }

        public IEnumerable<string> InspectionNames => _inspectionNames;

        public bool IsIgnored(string inspectionName)
        {
            return _inspectionNames.Contains(inspectionName);
        }

        public override string ToString()
        {
            return $"Ignored inspections: {string.Join(", ", _inspectionNames)}";
        }
    }
}
