using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Parsing.Annotations
{
    public abstract class FixedAttributeValueAnnotationBase : AnnotationBase, IAttributeAnnotation
    {
        private readonly string _attribute;
        private readonly IReadOnlyList<string> _attributeValues;

        protected FixedAttributeValueAnnotationBase(string name, AnnotationTarget target, string attribute, IEnumerable<string> attributeValues, bool allowMultiple = false, IReadOnlyList<ComponentType> incompatibleComponentTypes = null)
            : base(name, target, allowMultiple: allowMultiple)
        {
            // IEnumerable makes specifying the compile-time constant list easier on us
            _attributeValues = attributeValues.ToList();
            _attribute = attribute;
        }

        public override IReadOnlyList<ComponentType> IncompatibleComponentTypes { get; } = new[] { ComponentType.Document };

        public IReadOnlyList<string> AnnotationToAttributeValues(IReadOnlyList<string> annotationValues)
        {
            return _attributeValues;
        }

        public string Attribute(IReadOnlyList<string> annotationValues)
        {
            return _attribute;
        }

        public IReadOnlyList<string> AttributeToAnnotationValues(IReadOnlyList<string> attributeValues)
        {
            // annotation values must not be specified, because attribute values are fixed in the first place
            return new List<string>();
        }

        public bool MatchesAttributeDefinition(string attributeName, IReadOnlyList<string> attributeValues)
        {
            return _attribute == attributeName && _attributeValues.SequenceEqual(attributeValues);
        }
    }
}