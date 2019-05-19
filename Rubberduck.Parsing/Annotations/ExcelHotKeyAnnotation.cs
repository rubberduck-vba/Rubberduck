using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    public sealed class ExcelHotKeyAnnotation : FlexibleAttributeValueAnnotationBase
    {
        public ExcelHotKeyAnnotation(AnnotationType annotationType, QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IEnumerable<string> attributeValues) :
            base(annotationType, qualifiedSelection, context, GetHotKeyAttributeValue(attributeValues))
        { }

        private static IEnumerable<string> GetHotKeyAttributeValue(IEnumerable<string> attributeValues) => 
            attributeValues.Take(1).Select(StripStringLiteralQuotes).Select(v => v[0] + @"\n14").ToList();

        private static string StripStringLiteralQuotes(string value) =>
            value.StartsWith("\"") && value.EndsWith("\"") && value.Length > 2
                ? value.Substring(1, value.Length - 2)
                : value;
    }
}
