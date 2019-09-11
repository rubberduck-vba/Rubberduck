using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Annotations
{
    public sealed class VBAParserAnnotationFactory : IAnnotationFactory
    {
        private readonly Dictionary<string, IAnnotation> _lookup = new Dictionary<string, IAnnotation>();
        private readonly IAnnotation unrecognized;

        public VBAParserAnnotationFactory(IEnumerable<IAnnotation> recognizedAnnotations) 
        {
            foreach (var annotation in recognizedAnnotations)
            {
                if (annotation is NotRecognizedAnnotation)
                {
                    unrecognized = annotation;
                }
                _lookup.Add(annotation.Name.ToLowerInvariant(), annotation);
            }
        }

        public IParseTreeAnnotation Create(VBAParser.AnnotationContext context, QualifiedSelection qualifiedSelection)
        {
            var annotationName = context.annotationName().GetText();
            if (_lookup.TryGetValue(annotationName.ToLowerInvariant(), out var annotation))
            {
                return new ParseTreeAnnotation(annotation, qualifiedSelection, context);
            }
            return new ParseTreeAnnotation(unrecognized, qualifiedSelection, context);
        }
    }
}
