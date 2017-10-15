using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for specifying a member's <c>VB_UserMemId</c> attribute value.
    /// </summary>
    public sealed class DefaultMemberAnnotation : AnnotationBase, IAttributeAnnotation
    {
        public DefaultMemberAnnotation(QualifiedSelection qualifiedSelection, IEnumerable<string> parameters)
            : base(AnnotationType.DefaultMember, qualifiedSelection)
        {
            Description = parameters.FirstOrDefault();
        }

        public string Description { get; }
        public string Attribute => "VB_UserMemId = 0";
    }
}