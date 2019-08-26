using Rubberduck.VBEditor;
using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for all annotations not recognized by RD.
    /// Since this is not actually an annotation, it has no valid target
    /// </summary>
    [Annotation("NotRecognized", 0)]
    public sealed class NotRecognizedAnnotation : AnnotationBase
    {
        public NotRecognizedAnnotation(
            QualifiedSelection qualifiedSelection,
            VBAParser.AnnotationContext context,
            IEnumerable<string> parameters)
            : base(qualifiedSelection, context)
        {}
    }
}