using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    public abstract class FlexibleAttributeValueAnnotationBase : AnnotationBase, IAttributeAnnotation
    {
        public string Attribute { get; }

        private readonly int _numberOfValues;

        protected FlexibleAttributeValueAnnotationBase(string name, AnnotationTarget target, string attribute, int numberOfValues)
            : base(name, target)
        {
            Attribute = attribute;
            _numberOfValues = numberOfValues;
        }

        public bool MatchesAttributeDefinition(string attributeName, IReadOnlyList<string> attributeValues)
        {
            return Attribute == attributeName && _numberOfValues == attributeValues.Count;
        }

        public virtual IReadOnlyList<string> AnnotationToAttributeValues(IReadOnlyList<string> annotationValues)
        {
            return annotationValues.Take(_numberOfValues).Select(v => v.EnQuote()).ToList();
        }

        public virtual IReadOnlyList<string> AttributeToAnnotationValues(IReadOnlyList<string> attributeValues)
        {
            return attributeValues.Take(_numberOfValues).Select(v => v.EnQuote()).ToList();
        }

        string IAttributeAnnotation.Attribute(IReadOnlyList<string> annotationValues)
        {
            return Attribute;
        }
    }
}