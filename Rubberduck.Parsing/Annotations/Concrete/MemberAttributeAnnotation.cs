using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor.SafeComWrappers;
using System.Linq;

namespace Rubberduck.Parsing.Annotations.Concrete
{
    /// <summary>
    /// @MemberAttribute annotation, indicates the presence of a hidden member attribute; allows specifying arbitrary VB_Attribute for members. Use the quick-fixes to "Rubberduck Opportunities" code inspections to synchronize annotations and attributes.
    /// </summary>
    /// <parameter name="VB_Attribute" type="Identifier">
    /// The literal identifier name of the member VB_Attribute.
    /// </parameter>
    /// <parameter name="Values" type="ParamArray">
    /// The comma-separated attribute values, as applicable.
    /// </parameter>
    /// <remarks>
    /// The @MemberAttribute annotation cannot be used at module level. This separate annotation disambiguates any potential scoping issues that present themselves when the same name is used for both scopes.
    /// This annotation may be used with module variable targets.
    /// </remarks>
    /// <example>
    /// <before>
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// Option Explicit
    ///
    /// '@MemberAttribute VB_Description, "Does something"
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </module>
    /// </before>
    /// <after>
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// Option Explicit
    ///
    /// '@MemberAttribute VB_Description, "Does something"
    /// Public Sub DoSomething()
    /// Attribute DoSomething.VB_Description = "Does something"
    /// End Sub
    /// ]]>
    /// </module>
    /// </after>
    /// </example>
    public class MemberAttributeAnnotation : FlexibleAttributeAnnotationBase
    {
        public MemberAttributeAnnotation()
            : base("MemberAttribute", AnnotationTarget.Member | AnnotationTarget.Variable, _argumentTypes, true)
        {
            _incompatibleComponentTypes = base.IncompatibleComponentTypes
                .Concat(new[] { ComponentType.Document })
                .Distinct().ToList();
        }

        private readonly IReadOnlyList<ComponentType> _incompatibleComponentTypes;
        public override IReadOnlyList<ComponentType> IncompatibleComponentTypes => _incompatibleComponentTypes;
            

        private static readonly AnnotationArgumentType[] _argumentTypes = new[]
        {
            AnnotationArgumentType.Attribute,
            AnnotationArgumentType.Text | AnnotationArgumentType.Number | AnnotationArgumentType.Boolean
        };
    }
}