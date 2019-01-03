using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    public abstract class AttributeAnnotationBase : AnnotationBase, IAttributeAnnotation
    {
        protected AttributeAnnotationBase(AnnotationType annotationType, QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IReadOnlyList<string> attributeValues) 
        :base(annotationType, qualifiedSelection, context)
        {
            AttributeValues = attributeValues ?? new List<string>();
        }

        public abstract string Attribute { get; }
        public IReadOnlyList<string> AttributeValues { get; }
    }
}