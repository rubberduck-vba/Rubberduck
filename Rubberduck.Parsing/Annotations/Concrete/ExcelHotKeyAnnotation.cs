using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
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
