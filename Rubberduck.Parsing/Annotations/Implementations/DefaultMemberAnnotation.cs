using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for specifying a member's <c>VB_UserMemId</c> attribute value.
    /// </summary>
    /// 
    [Annotation("DefaultMember", AnnotationTarget.Member)]
    [FixedAttributeValueAnnotation("VB_UserMemId", "0")]
    public sealed class DefaultMemberAnnotation : FixedAttributeValueAnnotationBase
    {
        public DefaultMemberAnnotation(QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IEnumerable<string> parameters)
            : base(qualifiedSelection, context)
        {
            Description = parameters?.FirstOrDefault() ?? string.Empty;
        }

        public string Description { get; }
    }
}