using Rubberduck.VBEditor;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Annotations
{
    public sealed class TestMethodAnnotation : AnnotationBase
    {
        public TestMethodAnnotation(
            QualifiedSelection qualifiedSelection,
            IEnumerable<string> parameters)
            : base(AnnotationType.TestMethod, qualifiedSelection)
        {
        }
    }
}
