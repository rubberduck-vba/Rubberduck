using System.Collections.Generic;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Marks a module that Smart Indenter ignores.
    /// </summary>
    public sealed class NoIndentAnnotation : AnnotationBase
    {
        public NoIndentAnnotation(QualifiedSelection qualifiedSelection, IEnumerable<string> parameters)
            : base(AnnotationType.NoIndent, qualifiedSelection)
        {
        }
    }
}
