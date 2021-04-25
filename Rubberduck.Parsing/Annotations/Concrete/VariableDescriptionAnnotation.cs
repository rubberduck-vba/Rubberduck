using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Parsing.Annotations.Concrete
{
    /// <summary>
    /// @VariableDescription annotation, indicates the presence of a VB_VarDescription attribute value providing a docstring for a module-level variable or constant (not local variables). Use the quick-fixes to "Rubberduck Opportunities" code inspections to synchronize annotations and attributes.
    /// </summary>
    /// <parameter name="DocString" type="Text">
    /// This string literal parameter does not support expressions and/or multiline inputs. The string literal is used as-is as the value of the hidden member attribute.
    /// </parameter>
    /// <remarks>
    /// The @VariableDescription annotation complements the @description annotation, which can be applied to methods. Having separate annotations for variables and members disambiguates any potential scoping issues presenting themselves when the same name is used for both scopes.
    /// </remarks>
    /// <example>
    /// <before>
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// '@VariableDescription("Exposes a read/write value.")
    /// Public SomeValue As Long
    ///
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
    /// '@VariableDescription("Exposes a read/write value.")
    /// Public SomeValue As Long
    /// Attribute SomeValue.VB_VarDescription = "Exposes a read/write value."
    ///
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </module>
    /// </after>
    /// </example>
    public class VariableDescriptionAnnotation : DescriptionAttributeAnnotationBase
    {
        public VariableDescriptionAnnotation()
            : base("VariableDescription", AnnotationTarget.Variable, "VB_VarDescription")
        {}

        // override incompatibility for Document module to allow RD docstrings without the corresponding VB_Attribute.
        public override IReadOnlyList<ComponentType> IncompatibleComponentTypes { get; } = new ComponentType[] { };
    }
}