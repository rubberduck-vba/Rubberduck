using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using Rubberduck.Resources.Registration;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// @Enumerator annotation, indicates that the member should have a VB_UserMemId attribute value (-4) making it the enumerator-provider member of that class, enabling 'For Each' iteration of custom collections. Use the quick-fixes to "Rubberduck Opportunities" code inspections to synchronize annotations and attributes.
    /// </summary>
    /// <example>
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// Option Explicit
    /// Private InternalState As VBA.Collection
    ///
    /// @Enumerator
    /// Public Property Get NewEnum() As IUnknown
    ///     Set NewEnum = InternalState.[_NewEnum]
    /// End Sub
    /// 
    /// Private Sub Class_Initialize()
    ///     Set InternalState = New VBA.Collection
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    public sealed class EnumeratorMemberAnnotation : FixedAttributeValueAnnotationBase
    {
        public EnumeratorMemberAnnotation()
            : base("Enumerator", AnnotationTarget.Member, "VB_UserMemId", new[] { WellKnownDispIds.NewEnum.ToString() })
        {}
    }
}
