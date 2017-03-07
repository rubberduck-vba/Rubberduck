using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    public sealed class IgnoreModuleAnnotation : AnnotationBase
    {
        private readonly IEnumerable<string> _inspectionNames;

        public IgnoreModuleAnnotation(QualifiedSelection qualifiedSelection, IEnumerable<string> parameters)
            : base(AnnotationType.IgnoreModule, qualifiedSelection)
        {
            _inspectionNames = parameters;
        }

        public IEnumerable<string> InspectionNames
        {
            get
            {
                return _inspectionNames;
            }
        }

        public bool IsIgnored(string inspectionName)
        {
            return !_inspectionNames.Any() || _inspectionNames.Contains(inspectionName);
        }

        public override string ToString()
        {
            return string.Format("Ignored inspections: {0}", string.Join(", ", _inspectionNames));
        }
    }
}