using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for specifying a module's <c>VB_PredeclaredId</c> attribute.
    /// </summary>
    [Annotation("PredeclaredId", AnnotationTarget.Module)]
    [FixedAttributeValueAnnotation("VB_PredeclaredId", "True")]
    public sealed class PredeclaredIdAnnotation : FixedAttributeValueAnnotationBase
    {
        public PredeclaredIdAnnotation(QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IEnumerable<string> parameters)
            : base(qualifiedSelection, context)
        {}
    }
}