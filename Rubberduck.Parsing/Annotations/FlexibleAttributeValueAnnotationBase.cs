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

        protected FlexibleAttributeValueAnnotationBase(AnnotationType annotationType, QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IEnumerable<string> attributeValues)
        :base(annotationType, qualifiedSelection, context)
        {
            var flexibleAttributeValueInfo = FlexibleAttributeValueInfo(annotationType);

            Attribute = flexibleAttributeValueInfo.attribute;
            AttributeValues = attributeValues?.Take(flexibleAttributeValueInfo.numberOfValues).ToList() ?? new List<string>();
        }

        private static (string attribute, int numberOfValues) FlexibleAttributeValueInfo(AnnotationType annotationType)
        {
            var type = annotationType.GetType();
            var name = Enum.GetName(type, annotationType);
            var flexibleAttributeValueAttributes = type.GetField(name).GetCustomAttributes(false)
                .OfType<FlexibleAttributeValueAnnotationAttribute>().ToList();

            var attribute = flexibleAttributeValueAttributes.FirstOrDefault();

            if (attribute == null)
            {
                return ("", 0);
            }

            return (attribute.AttributeName, attribute.NumberOfParameters);
        }
    }
}