using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Parsing.Annotations.Concrete
{
    /// <summary>
    /// @ExcelHotkey annotation, indicates the presence of a VB_ProcData.VB_Invoke_Func metadata attribute value that maps a hotkey to a standard module procedure ("macro"). Use the quick-fixes to "Rubberduck Opportunities" code inspections to synchronize annotations and attributes.
    /// </summary>
    /// <parameter name="Key" type="Text">
    /// A single-letter string argument maps the hotkey. If the letter is in UPPER CASE, the hotkey is Ctrl+Shift+letter; if the letter is lower case, the hotkey is Ctrl+letter. Avoid remapping commonly used keyboard shortcuts!
    /// </parameter>
    /// <remarks>
    /// Members with this annotation are ignored by the ProcedureNotUsed inspection. Use the @EntryPoint annotation to similarly affect the ProcedureNotUsed inspection without mapping a hotkey or ignoring the inspection.
    /// </remarks>
    /// <example>
    /// <before>
    /// <module name="Module1" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    ///
    /// '@ExcelHotkey "D"
    /// Public Sub DoSomething()
    ///     '...
    /// End Sub
    /// ]]>
    /// </module>
    /// </before>
    /// <after>
    /// <module name="Module1" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    ///
    /// '@ExcelHotkey "D"
    /// Public Sub DoSomething()
    /// Attribute DoSomething.VB_ProcData.VB_Invoke_Func = "D\n14"
    ///     '...
    /// End Sub
    /// ]]>
    /// </module>
    /// </after>
    /// </example>
    public sealed class ExcelHotKeyAnnotation : FlexibleAttributeValueAnnotationBase
    {
        public ExcelHotKeyAnnotation()
            : base("ExcelHotkey", AnnotationTarget.Member, "VB_ProcData.VB_Invoke_Func", 1, new[] { AnnotationArgumentType.Text}) 
        {
        }

        public override ComponentType? RequiredComponentType => ComponentType.StandardModule;

        public override IReadOnlyList<string> AnnotationToAttributeValues(IReadOnlyList<string> annotationValues) =>
            annotationValues.Take(1).Select(v => (v.UnQuote()[0] + @"\n14").EnQuote()).ToList();
        
        public override IReadOnlyList<string> AttributeToAnnotationValues(IReadOnlyList<string> attributeValues) =>        
            attributeValues.Select(keySpec => keySpec.UnQuote().Substring(0, 1)).ToList();
       
    }
}
