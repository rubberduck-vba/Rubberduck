using Rubberduck.VBEditor;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Marks a method that the test engine will execute as a unit test.
    /// </summary>
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
