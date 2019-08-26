using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    [Annotation("ExcelHotkey", AnnotationTarget.Member)]
    [FlexibleAttributeValueAnnotation("VB_ProcData.VB_Invoke_Func", 1, true)]
    public sealed class ExcelHotKeyAnnotation : FlexibleAttributeValueAnnotationBase
    {
        public ExcelHotKeyAnnotation(QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IEnumerable<string> annotationParameterValues) :
            base(qualifiedSelection, context, GetHotKeyAttributeValue(annotationParameterValues))
        { }
        
        private static IEnumerable<string> GetHotKeyAttributeValue(IEnumerable<string> parameters) => 
            parameters.Take(1).Select(v => v.UnQuote()[0] + @"\n14".EnQuote()).ToList();
        
        public static IEnumerable<string> TransformToAnnotationValues(IEnumerable<string> attributeValues) =>
            attributeValues.Select(keySpec => keySpec.UnQuote().Substring(0, 1));
    }
}
