using Rubberduck.VBEditor;
using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Marks a method that the test engine will execute before executing the first unit test in a test module.
    /// </summary>
    public sealed class ModuleInitializeAnnotation : AnnotationBase
    {
        public ModuleInitializeAnnotation(
            QualifiedSelection qualifiedSelection,
            VBAParser.AnnotationContext context,
            IEnumerable<string> parameters)
            : base(AnnotationType.ModuleInitialize, qualifiedSelection, context)
        {
        }
    }
}
