using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Annotations
{
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

        public IEnumerable<string> InspectionNames
        {
            get
            {
                return _inspectionNames;
            }
        }

        public bool IsIgnored(string inspectionName)
        {
            return _inspectionNames.Contains(inspectionName);
        }

        public override string ToString()
        {
            return string.Format("Ignored inspections: {0}", string.Join(", ", _inspectionNames));
        }
    }
}
