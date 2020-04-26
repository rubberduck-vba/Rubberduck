using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// @ExcelHotkey annotation, uses a VB_ProcData.VB_Invoke_Func metadata attribute to map a hotkey to a standard module procedure. Use the quick-fixes to "Rubberduck Opportunities" code inspections to synchronize annotations and attributes.
    /// </summary>
    /// <parameter name="Key" type="String*1">
    /// A single-letter string argument maps the hotkey. If the letter is in UPPER CASE, the hotkey is Ctrl+Shift+letter; if the letter is lower case, the hotkey is Ctrl+letter. Avoid remapping commonly used keyboard shortcuts!
    /// </parameter>
    /// <example>
    /// <module name="Module1" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    ///
    /// @ExcelHotkey("D")
    /// Public Sub DoSomething()
    ///     '...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    public sealed class ExcelHotKeyAnnotation : FlexibleAttributeValueAnnotationBase
    {
        public ExcelHotKeyAnnotation()
            : base("ExcelHotkey", AnnotationTarget.Member, "VB_ProcData.VB_Invoke_Func", 1)
        { }

        public override IReadOnlyList<string> AnnotationToAttributeValues(IReadOnlyList<string> annotationValues)
        {
            return annotationValues.Take(1).Select(v => (v.UnQuote()[0] + @"\n14").EnQuote()).ToList();
        }

        public override IReadOnlyList<string> AttributeToAnnotationValues(IReadOnlyList<string> attributeValues)
        {
            return attributeValues.Select(keySpec => keySpec.UnQuote().Substring(0, 1)).ToList();
        }
    }
}
