using Rubberduck.VBEditor;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Annotations
{
    public sealed class TestCleanupAnnotation : AnnotationBase
    {
        public TestCleanupAnnotation(
            QualifiedSelection qualifiedSelection,
            IEnumerable<string> parameters)
            : base(AnnotationType.TestCleanup, qualifiedSelection)
        {
        }
    }
}
