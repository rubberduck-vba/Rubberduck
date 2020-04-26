using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// @ModuleDescription annotation, uses the VB_Description module attribute to provide a docstring for a module. Use the quick-fixes to "Rubberduck Opportunities" code inspections to synchronize annotations and attributes.
    /// </summary>
    /// <parameter>
    /// This annotation takes a single string literal parameter that does not support expressions and/or multiline inputs. The string literal is used as-is as the value of the hidden module attribute.
    /// </parameter>
    /// <remarks>
    /// The @Description annotation cannot be used at module level. This separate annotation disambiguates any potential scoping issues that present themselves when the same name is used for both scopes.
    /// </remarks>
    /// <example>
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// '@ModuleDescription("Represents an object responsible for doing something.")
    /// Option Explicit
    ///
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    public sealed class ModuleDescriptionAnnotation : DescriptionAttributeAnnotationBase
    {
        public ModuleDescriptionAnnotation()
            : base("ModuleDescription", AnnotationTarget.Module, "VB_Description", 1)
        {}
    }
}