using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Annotations;

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
    }
}