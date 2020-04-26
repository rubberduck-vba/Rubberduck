using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// @Exposed annotation, uses the VB_Exposed module attribute to make a class visible to a referencing project (classes are otherwise private). Use the quick-fixes to "Rubberduck Opportunities" code inspections to synchronize annotations and attributes.
    /// </summary>
    /// <example>
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// '@Exposed
    /// Option Explicit
    ///
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    public sealed class ExposedModuleAnnotation : FixedAttributeValueAnnotationBase
    {
        public ExposedModuleAnnotation()
            : base("Exposed", AnnotationTarget.Module, "VB_Exposed", new[] { "True" })
        {}
    }
}