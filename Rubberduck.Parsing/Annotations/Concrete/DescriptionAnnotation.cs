using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Parsing.Annotations.Concrete
{
    /// <summary>
    /// @Description annotation, indicates that the member should have a VB_Description attribute to provide a docstring. Use the quick-fixes to "Rubberduck Opportunities" code inspections to synchronize annotations and attributes.
    /// </summary>
    /// <parameter name="DocString" type="Text">
    /// This string literal parameter does not support expressions and/or multiline inputs. The string literal is used as-is as the value of the hidden member attribute.
    /// </parameter>
    /// <remarks>
    /// This documentation string appears in the VBE's own Object Browser, as well as in various Rubberduck UI elements.
    /// </remarks>
    /// <example>
    /// <before>
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// Option Explicit
    ///
    /// '@Description("Does something")
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
    /// '@Description("Does something")
    /// Public Sub DoSomething()
    /// Attribute DoSomething.VB_Description = "Does something"
    /// End Sub
    /// ]]>
    /// </module>
    /// </after>
    /// </example>
    public sealed class DescriptionAnnotation : DescriptionAttributeAnnotationBase
    {
        public DescriptionAnnotation()
            : base("Description", AnnotationTarget.Member, "VB_Description")
        {}

        // override incompatibility for Document module to allow RD docstrings without the corresponding VB_Attribute.
        public override IReadOnlyList<ComponentType> IncompatibleComponentTypes { get; } = new ComponentType[] { };
    }
}