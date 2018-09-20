using Rubberduck.VBEditor;
using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Marks a method that the test engine will execute after all unit tests in a test module have executed.
    /// </summary>
    public sealed class ModuleCleanupAnnotation : AnnotationBase
    {
        public ModuleCleanupAnnotation(
            QualifiedSelection qualifiedSelection,
            VBAParser.AnnotationContext context,
            IEnumerable<string> parameters)
            : base(AnnotationType.ModuleCleanup, qualifiedSelection, context)
        {
        }
    }
}
