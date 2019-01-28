using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    public abstract class FlexibleAttributeAnnotationBase : AnnotationBase, IAttributeAnnotation
    {
        protected FlexibleAttributeAnnotationBase(AnnotationType annotationType, QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IReadOnlyList<string> parameters) 
            :base(annotationType, qualifiedSelection, context)
        {
            Attribute = parameters?.FirstOrDefault() ?? string.Empty;
            AttributeValues = parameters?.Skip(1).ToList() ?? new List<string>();
        }

        public string Attribute { get; }
        public IReadOnlyList<string> AttributeValues { get; }
    }
}