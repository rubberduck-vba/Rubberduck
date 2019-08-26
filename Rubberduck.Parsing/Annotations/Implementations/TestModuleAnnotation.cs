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
    [Annotation("TestModule", AnnotationTarget.Module)]
    public sealed class TestModuleAnnotation : AnnotationBase
    {
        // TODO investigate unused parameters argument. Possibly needed to match signature for construction through VBAParserAnnotationFactory?!
        public TestModuleAnnotation(
            QualifiedSelection qualifiedSelection,
            VBAParser.AnnotationContext context,
            IEnumerable<string> parameters)
            : base(qualifiedSelection, context)
        {
        }
    }
}
