using Rubberduck.Resources.Registration;
using Rubberduck.Parsing.Annotations;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers;
using System.Linq;

namespace Rubberduck.Parsing.Annotations.Concrete
{
    /// <summary>
    /// @DefaultMember annotation, indicates that the member should have a VB_UserMemId attribute value (0) making it the default member of that class. Use the quick-fixes to "Rubberduck Opportunities" code inspections to synchronize annotations and attributes.
    /// </summary>
    /// <example>
    /// <before>
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// Option Explicit
    /// Private InternalState As VBA.Collection
    ///
    /// '@DefaultMember
    /// Public Property Get Item(ByVal Index As Variant) As Variant
    ///     Item = InternalState(Index)
    /// End Sub
    /// 
    /// 'if the default member is a property, only the Get accessor needs the attribute/annotation.
    /// Public Property Let Item(ByVal Index As Variant, ByVal Value As Variant)
    ///     InternalState(Index) = Value    
    /// End Sub
    /// 
    /// Private Sub Class_Initialize()
    ///     Set InternalState = New VBA.Collection
    /// End Sub
    /// ]]>
    /// </module>
    /// </before>
    /// <after>
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// Option Explicit
    /// Private InternalState As VBA.Collection
    ///
    /// '@DefaultMember
    /// Public Property Get Item(ByVal Index As Variant) As Variant
    /// Attribute Item.VB_UserMemId = 0
    ///     Item = InternalState(Index)
    /// End Sub
    /// 
    /// 'if the default member is a property, only the Get accessor needs the attribute/annotation.
    /// Public Property Let Item(ByVal Index As Variant, ByVal Value As Variant)
    ///     InternalState(Index) = Value    
    /// End Sub
    /// 
    /// Private Sub Class_Initialize()
    ///     Set InternalState = New VBA.Collection
    /// End Sub
    /// ]]>
    /// </module>
    /// </after>
    /// </example>
    public sealed class DefaultMemberAnnotation : FixedAttributeValueAnnotationBase
    {
        public DefaultMemberAnnotation()
            : base("DefaultMember", AnnotationTarget.Member, "VB_UserMemId", new[] { WellKnownDispIds.Value.ToString() })
        {}

        public override IReadOnlyList<ComponentType> IncompatibleComponentTypes =>
            base.IncompatibleComponentTypes.Concat(new[] { ComponentType.StandardModule }).Distinct().ToList();
    }
}
