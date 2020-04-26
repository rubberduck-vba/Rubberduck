using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// @ModuleAttribute annotation, allows specifying arbitrary VB_Attribute for modules. Use the quick-fixes to "Rubberduck Opportunities" code inspections to synchronize annotations and attributes.
    /// </summary>
    /// <parameter name="VB_Attribute" type="Identifier">
    /// The literal identifier name of the member VB_Attribute.
    /// </parameter>
    /// <parameter name="Values" type="ParamArray">
    /// The comma-separated attribute values, as applicable.
    /// </parameter>
    /// <example>
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// '@ModuleAttribute VB_Ext_Key, "Key", "Value"
    /// Option Explicit
    ///
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    public class ModuleAttributeAnnotation : FlexibleAttributeAnnotationBase
    {
        public ModuleAttributeAnnotation() 
        : base("ModuleAttribute", AnnotationTarget.Module)
        {}
    }
}