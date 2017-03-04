using Rubberduck.VBEditor;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Annotations
{
    public sealed class TestModuleAnnotation : AnnotationBase
    {
        public TestModuleAnnotation(
            QualifiedSelection qualifiedSelection,
            IEnumerable<string> parameters)
            : base(AnnotationType.TestModule, qualifiedSelection)
        {
        }
    }
}
