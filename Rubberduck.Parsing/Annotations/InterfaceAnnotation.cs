using System.Collections.Generic;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    public sealed class InterfaceAnnotation : AnnotationBase
    {
        public InterfaceAnnotation(QualifiedSelection qualifiedSelection, IEnumerable<string> parameters)
            : base(AnnotationType.Interface, qualifiedSelection)
        {
        }
    }
}