using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Annotations;

namespace Rubberduck.Parsing.Annotations.Concrete
{
    /// <summary>
    /// @Exposed annotation, indicates the presence of a VB_Exposed module attribute value (True) to make a class visible to a referencing project (classes are otherwise private by default). Use the quick-fixes to "Rubberduck Opportunities" code inspections to synchronize annotations and attributes.
    /// </summary>
    /// <example>
    /// <before>
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// '@Exposed
    /// Option Explicit
    ///
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </module>
    /// </before>
    /// <after>
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// Attribute VB_Exposed = True
    /// '@Exposed
    /// Option Explicit
    ///
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </module>
    /// </after>
    /// </example>
    public sealed class ExposedModuleAnnotation : FixedAttributeValueAnnotationBase
    {
        public ExposedModuleAnnotation()
            : base("Exposed", AnnotationTarget.Module, "VB_Exposed", new[] { "True" })
        {}
    }
}