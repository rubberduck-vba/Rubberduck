using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// @Description annotation, uses the VB_Description member attribute to provide a docstring for a module member. Use the quick-fixes to "Rubberduck Opportunities" code inspections to synchronize annotations and attributes.
    /// </summary>
    /// <parameter>
    /// This annotation takes a single string literal parameter that does not support expressions and/or multiline inputs. The string literal is used as-is as the value of the hidden member attribute.
    /// </parameter>
    /// <example>
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// Option Explicit
    ///
    /// @Description("Does something")
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    public sealed class DescriptionAnnotation : DescriptionAttributeAnnotationBase
    {
        public DescriptionAnnotation()
            : base("Description", AnnotationTarget.Member, "VB_Description", 1)
        {}
    }
}