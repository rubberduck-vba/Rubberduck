using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Annotations
{
    public static class AttributeExtensions
    {
        private static readonly IDictionary<AttributeAnnotationAttribute, AnnotationType> AttributeAnnotationTypes = 
            typeof(AnnotationType)
                .GetFields()
                .Where(field => field
                    .GetCustomAttributes(typeof(AttributeAnnotationAttribute), true)
                    .Any())
                .Select(field => new
                {
                    AnnotationAttributes = field
                        .GetCustomAttributes(typeof(AttributeAnnotationAttribute), true)
                        .Cast<AttributeAnnotationAttribute>(),
                    AnnotationType = (AnnotationType)Enum.Parse(typeof(AnnotationType), field.Name)
                })
                .ToDictionary(a => a.AnnotationAttributes.FirstOrDefault(), a => a.AnnotationType);

        /// <summary>
        /// Gets the <see cref="AnnotationType"/>, if any, associated with the context.
        /// </summary>
        /// <returns>Returns <c>null</c> if no association is found.</returns>
        public static AnnotationType? AnnotationType(this VBAParser.AttributeStmtContext context)
        {
            var name = context.attributeName().Stop.Text;
            var value = context.attributeValue().FirstOrDefault()?.GetText();

            var key = AttributeAnnotationTypes.Keys
                                      // EndsWith the name, so we can match "VB_Description" to "DoSomething.VB_Description"
                .SingleOrDefault(k => k.AttributeName.EndsWith(name) && (k.AttributeValue?.Equals(value, StringComparison.OrdinalIgnoreCase) ?? true));

            AnnotationType result;
            return key != null && AttributeAnnotationTypes.TryGetValue(key, out result)
                ? (AnnotationType?) result
                : null;
        }
    }
}