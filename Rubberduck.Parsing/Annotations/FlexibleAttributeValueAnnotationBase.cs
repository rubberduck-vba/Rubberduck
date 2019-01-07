using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    public abstract class FlexibleAttributeValueAnnotationBase : AnnotationBase, IAttributeAnnotation
    {
        protected FlexibleAttributeValueAnnotationBase(AnnotationType annotationType, QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IEnumerable<string> parameters)
        :base(annotationType, qualifiedSelection, context)
        {
            var flexibleAttributeValueInfo = FlexibleAttributeValueInfo(annotationType);

            if (flexibleAttributeValueInfo == null)
            {
                Attribute = string.Empty;
                AttributeValues = new List<string>();
                return;
            }

            Attribute = flexibleAttributeValueInfo.Value.attribute;
            AttributeValues = parameters?.Take(flexibleAttributeValueInfo.Value.numberOfValues).ToList() ?? new List<string>();
        }

        public string Attribute { get; }
        public IReadOnlyList<string> AttributeValues { get; }

        private static (string attribute, int numberOfValues)? FlexibleAttributeValueInfo(AnnotationType annotationType)
        {
            var type = annotationType.GetType();
            var name = Enum.GetName(type, annotationType);
            var flexibleAttributeValueAttributes = type.GetField(name).GetCustomAttributes(false)
                .OfType<FlexibleAttributeValueAnnotationAttribute>().ToList();

            var attribute = flexibleAttributeValueAttributes.FirstOrDefault();

            if (attribute == null)
            {
                return null;
            }

            return (attribute.AttributeName, attribute.NumberOfParameters);
        }
    }
}