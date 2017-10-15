using Rubberduck.VBEditor;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Annotations
{
    public sealed class TestInitializeAnnotation : AnnotationBase
    {
        public TestInitializeAnnotation(
            QualifiedSelection qualifiedSelection,
            IEnumerable<string> parameters)
            : base(AnnotationType.TestInitialize, qualifiedSelection)
        {
        }
    }
}
