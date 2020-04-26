using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// @VariableDescription annotation, uses the VB_VarDescription attribute to provide a docstring for a module-level variable or constant. Use the quick-fixes to "Rubberduck Opportunities" code inspections to synchronize annotations and attributes.
    /// </summary>
    /// <parameter>
    /// This annotation takes no argument.
    /// </parameter>
    /// <remarks>
    /// The @Description annotation cannot be used at the module variable level. This separate annotation disambiguates any potential scoping issues that present themselves when the same name is used for both scopes.
    /// </remarks>
    /// <example>
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
    /// </example>
    public class VariableDescriptionAnnotation : DescriptionAttributeAnnotationBase
    {
        public VariableDescriptionAnnotation()
            : base("VariableDescription", AnnotationTarget.Variable, "VB_VarDescription", 1)
        {}
    }
}