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
        public NoIndentAnnotation()
            : base("NoIndent", AnnotationTarget.Module)
        { }
    }
}
