using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    public abstract class FlexibleAttributeValueAnnotationBase : AnnotationBase, IAttributeAnnotation
    {
        public string Attribute { get; }
        public IReadOnlyList<string> AttributeValues { get; }

        protected FlexibleAttributeValueAnnotationBase(QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IEnumerable<string> attributeValues)
        :base(qualifiedSelection, context)
        {
            var flexibleAttributeValueInfo = FlexibleAttributeValueInfo(GetType());

            Attribute = flexibleAttributeValueInfo.attribute;
            AttributeValues = attributeValues?.Take(flexibleAttributeValueInfo.numberOfValues).ToList() ?? new List<string>();
        }

        private static (string attribute, int numberOfValues) FlexibleAttributeValueInfo(Type annotationType)
        {
            var attributeValueInfo = annotationType.GetCustomAttributes(false)
                .OfType<FlexibleAttributeValueAnnotationAttribute>()
                .SingleOrDefault();

            if (attributeValueInfo == null)
            {
                return ("", 0);
            }
            return (attributeValueInfo.AttributeName, attributeValueInfo.NumberOfParameters);
        }
    }
}