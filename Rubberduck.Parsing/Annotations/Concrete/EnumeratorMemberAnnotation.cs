using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using Rubberduck.Resources.Registration;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for specifying a member's <c>VB_UserMemId</c> attribute value.
    /// </summary>
    public sealed class EnumeratorMemberAnnotation : FixedAttributeValueAnnotationBase
    {
        public EnumeratorMemberAnnotation()
            : base("Enumerator", AnnotationTarget.Member, "VB_UserMemId", new[] { WellKnownDispIds.NewEnum.ToString() })
        {}
    }
}
