﻿using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    public abstract class FlexibleAttributeAnnotationBase : AnnotationBase, IAttributeAnnotation
    {
        protected FlexibleAttributeAnnotationBase(string name, AnnotationTarget target, bool allowMultiple = false)
            : base(name, target, allowMultiple)
        { }
        
        public IReadOnlyList<string> AnnotationToAttributeValues(IReadOnlyList<string> annotationValues)
        {
            // skip the attribute specification, which is taken from the annotationValues
            // also we MUST NOT adjust quotation of annotationValues here
            return annotationValues?.Skip(1).ToList();
        }

        public string Attribute(IReadOnlyList<string> annotationValues)
        {
            // The Attribute name is NEVER quoted, therefore unquote here
            return annotationValues.FirstOrDefault()?.UnQuote() ?? "";
        }

        public IReadOnlyList<string> AttributeToAnnotationValues(IReadOnlyList<string> attributeValues)
        {
            // Must not adjust quotation status
            return attributeValues;
        }

        public bool MatchesAttributeDefinition(string attributeName, IReadOnlyList<string> attributeValues)
        {
            // Implementers are the fallback. They must not return true here to avoid locking out more suitable candidates
            return false;
        }
    }
}