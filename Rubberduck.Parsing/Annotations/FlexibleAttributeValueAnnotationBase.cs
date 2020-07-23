using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;

namespace Rubberduck.Parsing.Annotations
{
    public abstract class FlexibleAttributeValueAnnotationBase : AnnotationBase, IAttributeAnnotation
    {
        private readonly string _attribute;
        private readonly int _numberOfValues;

        protected FlexibleAttributeValueAnnotationBase(string name, AnnotationTarget target, string attribute, int numberOfValues)
            : base(name, target, numberOfValues, numberOfValues)
        {
            _attribute = attribute;
            _numberOfValues = numberOfValues;
        }

        public bool MatchesAttributeDefinition(string attributeName, IReadOnlyList<string> attributeValues)
        {
            return _attribute == attributeName && _numberOfValues == attributeValues.Count;
        }

        public virtual IReadOnlyList<string> AnnotationToAttributeValues(IReadOnlyList<string> annotationValues)
        {
            return annotationValues.Take(_numberOfValues).Select(v => v.EnQuote()).ToList();
        }

        public virtual IReadOnlyList<string> AttributeToAnnotationValues(IReadOnlyList<string> attributeValues)
        {
            return attributeValues.Take(_numberOfValues).Select(v => v.EnQuote()).ToList();
        }

        public string Attribute(IReadOnlyList<string> annotationValues)
        {
            return _attribute;
        }
    }
}