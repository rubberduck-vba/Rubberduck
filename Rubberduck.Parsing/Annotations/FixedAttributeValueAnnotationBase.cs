using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    public abstract class FixedAttributeValueAnnotationBase : AnnotationBase, IAttributeAnnotation
    {
        protected FixedAttributeValueAnnotationBase(AnnotationType annotationType, QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context)
            : base(annotationType, qualifiedSelection, context)
        {
            var fixedAttributeValueInfo = FixedAttributeValueInfo(annotationType);

            Attribute = fixedAttributeValueInfo?.attribute ?? string.Empty;
            AttributeValues = fixedAttributeValueInfo?.attributeValues ?? new List<string>();
        }

        public string Attribute { get; }
        public IReadOnlyList<string> AttributeValues { get; }

        private static (string attribute, IReadOnlyList<string> attributeValues)? FixedAttributeValueInfo(AnnotationType annotationType)
        {
            var type = annotationType.GetType();
            var name = Enum.GetName(type, annotationType);
            var flexibleAttributeValueAttributes = type.GetField(name).GetCustomAttributes(false)
                .OfType<FixedAttributeValueAnnotationAttribute>().ToList();

            var attribute = flexibleAttributeValueAttributes.FirstOrDefault();

            if (attribute == null)
            {
                return null;
            }

            return (attribute.AttributeName, attribute.AttributeValues);
        }
    }
}