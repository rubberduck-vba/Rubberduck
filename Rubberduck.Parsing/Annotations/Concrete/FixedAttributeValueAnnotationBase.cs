using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    public abstract class FixedAttributeValueAnnotationBase : AnnotationBase, IAttributeAnnotation
    {
        private readonly string attribute;
        private readonly IReadOnlyList<string> attributeValues;

        protected FixedAttributeValueAnnotationBase(string name, AnnotationTarget target, string attribute, IEnumerable<string> attributeValues, bool allowMultiple = false)
            : base(name, target, allowMultiple)
        {
            // IEnumerable makes specifying the compile-time constant list easier on us
            this.attributeValues = attributeValues.ToList();
            this.attribute = attribute;
        }

        public IReadOnlyList<string> AnnotationToAttributeValues(IReadOnlyList<string> annotationValues)
        {
            return attributeValues;
        }

        public string Attribute(IReadOnlyList<string> annotationValues)
        {
            return attribute;
        }

        public IReadOnlyList<string> AttributeToAnnotationValues(IReadOnlyList<string> attributeValues)
        {
            // annotation values must not be specified, because attribute values are fixed in the first place
            return new List<string>();
        }

        public bool MatchesAttributeDefinition(string attributeName, IReadOnlyList<string> attributeValues)
        {
            return attribute == attributeName && this.attributeValues.SequenceEqual(attributeValues);
        }
    }
}