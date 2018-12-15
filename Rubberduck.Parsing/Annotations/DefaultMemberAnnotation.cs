using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for specifying a member's <c>VB_UserMemId</c> attribute value.
    /// </summary>
    public sealed class DefaultMemberAnnotation : AttributeAnnotationBase
    {
        public DefaultMemberAnnotation(QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IEnumerable<string> parameters)
            : base(AnnotationType.DefaultMember, qualifiedSelection, context, new List<string> { "0" })
        {
            Description = parameters?.FirstOrDefault() ?? string.Empty;
        }

        public string Description { get; }
        public override string Attribute => "VB_UserMemId";
    }
}