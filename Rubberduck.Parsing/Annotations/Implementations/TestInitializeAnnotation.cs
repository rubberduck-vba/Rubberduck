using Rubberduck.VBEditor;
using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Marks a method that the test engine will execute before executing each unit test in a test module.
    /// </summary>
    [Annotation("TestInitialize", AnnotationTarget.Member)]
    public sealed class TestInitializeAnnotation : AnnotationBase
    {
        public TestInitializeAnnotation(
            QualifiedSelection qualifiedSelection,
            VBAParser.AnnotationContext context,
            IEnumerable<string> parameters)
            : base(qualifiedSelection, context)
        {
        }
    }
}
