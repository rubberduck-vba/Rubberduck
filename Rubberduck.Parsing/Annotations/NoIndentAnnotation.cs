using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Marks a module that Smart Indenter ignores.
    /// </summary>
    public sealed class NoIndentAnnotation : AnnotationBase
    {
        public NoIndentAnnotation(QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IEnumerable<string> parameters)
            : base(AnnotationType.NoIndent, qualifiedSelection, context)
        {
        }
    }
}
