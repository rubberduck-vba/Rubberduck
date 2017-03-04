using System.Collections.Generic;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    public sealed class NoIndentAnnotation : AnnotationBase
    {
        public NoIndentAnnotation(QualifiedSelection qualifiedSelection, IEnumerable<string> parameters)
            : base(AnnotationType.NoIndent, qualifiedSelection)
        {
        }
    }
}
