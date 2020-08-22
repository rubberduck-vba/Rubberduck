using Rubberduck.Resources.Registration;
using Rubberduck.Parsing.Annotations;

namespace Rubberduck.Parsing.Annotations.Concrete
{
    /// <summary>
    /// @Enumerator annotation, indicates that the member should have a VB_UserMemId attribute value (-4) making it the enumerator-provider member of that class, enabling 'For Each' iteration of custom collections. Use the quick-fixes to "Rubberduck Opportunities" code inspections to synchronize annotations and attributes.
    /// </summary>
    /// <example>
    /// <module name="Class1" type="Class Module">
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// Private InternalState As VBA.Collection
    ///
    /// '@Enumerator
    /// Public Property Get NewEnum() As IUnknown
    ///     Set NewEnum = InternalState.[_NewEnum]
    /// End Sub
    /// 
    /// Private Sub Class_Initialize()
    ///     Set InternalState = New VBA.Collection
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// Private InternalState As VBA.Collection
    ///
    /// '@Enumerator
    /// Public Property Get NewEnum() As IUnknown
    /// Attribute NewEnum.VB_UserMemId = -4
    ///     Set NewEnum = InternalState.[_NewEnum]
    /// End Sub
    /// 
    /// Private Sub Class_Initialize()
    ///     Set InternalState = New VBA.Collection
    /// End Sub
    /// ]]>
    /// </after>
    /// </module>
    /// </example>
    public sealed class EnumeratorMemberAnnotation : FixedAttributeValueAnnotationBase
    {
        public EnumeratorMemberAnnotation()
            : base("Enumerator", AnnotationTarget.Member, "VB_UserMemId", new[] { WellKnownDispIds.NewEnum.ToString() })
        {}
    }
}
