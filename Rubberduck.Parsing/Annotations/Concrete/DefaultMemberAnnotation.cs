using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using Rubberduck.Resources.Registration;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for specifying a member's <c>VB_UserMemId</c> attribute value.
    /// </summary>
    public sealed class DefaultMemberAnnotation : FixedAttributeValueAnnotationBase
    {
        public DefaultMemberAnnotation()
            : base("DefaultMember", AnnotationTarget.Member, "VB_UserMemId", new[] { WellKnownDispIds.Value.ToString() })
        {
        }
    }
}
