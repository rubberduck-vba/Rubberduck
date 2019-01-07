using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for specifying a member's <c>VB_UserMemId</c> attribute value.
    /// </summary>
    public sealed class EnumeratorMemberAnnotation : FixedAttributeValueAnnotationBase
    {
        public EnumeratorMemberAnnotation(QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IEnumerable<string> parameters)
            : base(AnnotationType.Enumerator, qualifiedSelection, context)
        {}
    }
}