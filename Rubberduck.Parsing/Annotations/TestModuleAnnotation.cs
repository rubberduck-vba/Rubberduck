using Rubberduck.VBEditor;
using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Marks a module that the test engine treat as a test module.
    /// </summary>
    /// <remarks>
    /// Unit test discovery only inspects modules with a <c>@TestModule</c> annotation.
    /// </remarks>
    public sealed class TestModuleAnnotation : AnnotationBase
    {
        public TestModuleAnnotation(
            QualifiedSelection qualifiedSelection,
            VBAParser.AnnotationContext context,
            IEnumerable<string> parameters)
            : base(AnnotationType.TestModule, qualifiedSelection, context)
        {
        }
    }
}
