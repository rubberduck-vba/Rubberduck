using Rubberduck.VBEditor;
using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for all annotations not recognized by RD.
    /// </summary>
    public sealed class NotRecognizedAnnotation : AnnotationBase
    {
        public NotRecognizedAnnotation(
            QualifiedSelection qualifiedSelection,
            VBAParser.AnnotationContext context,
            IEnumerable<string> parameters)
            : base(AnnotationType.NotRecognized, qualifiedSelection, context)
        {}
    }
}