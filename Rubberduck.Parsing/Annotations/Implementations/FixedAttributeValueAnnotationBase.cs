using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    public abstract class FixedAttributeValueAnnotationBase : AnnotationBase, IAttributeAnnotation
    {
        protected FixedAttributeValueAnnotationBase(QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context)
            : base(qualifiedSelection, context)
        {
            var fixedAttributeValueInfo = FixedAttributeValueInfo(GetType());

            Attribute = fixedAttributeValueInfo.attribute;
            AttributeValues = fixedAttributeValueInfo.attributeValues;
        }

        public string Attribute { get; }
        public IReadOnlyList<string> AttributeValues { get; }

        private static (string attribute, IReadOnlyList<string> attributeValues) FixedAttributeValueInfo(Type annotationType)
        {
            var attributeValueInfo = annotationType.GetCustomAttributes(false)
                .OfType<FixedAttributeValueAnnotationAttribute>()
                .SingleOrDefault();
            if (attributeValueInfo == null)
            {
                return ("", new List<string>());
            }

            return (attributeValueInfo.AttributeName, attributeValueInfo.AttributeValues);
        }
    }
}