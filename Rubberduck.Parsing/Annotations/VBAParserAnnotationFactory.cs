using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Annotations
{
    public sealed class VBAParserAnnotationFactory : IAnnotationFactory
    {
        private readonly Dictionary<string, Type> _creators = new Dictionary<string, Type>();

        public VBAParserAnnotationFactory(IEnumerable<Type> recognizedAnnotationTypes) 
        {
            foreach (var annotationType in recognizedAnnotationTypes)
            {
                // Extract the static information about the annotation type from it's AnnotationAttribute
                var staticInfo = annotationType.GetCustomAttributes(false)
                    .OfType<AnnotationAttribute>()
                    .Single();
                _creators.Add(staticInfo.Name.ToUpperInvariant(), annotationType);
            }
        }

        public IAnnotation Create(VBAParser.AnnotationContext context, QualifiedSelection qualifiedSelection)
        {
            var annotationName = context.annotationName().GetText();
            var parameters = AnnotationParametersFromContext(context);
            return CreateAnnotation(annotationName, parameters, qualifiedSelection, context);
        }

        private static List<string> AnnotationParametersFromContext(VBAParser.AnnotationContext context)
        {
            var parameters = new List<string>();
            var argList = context.annotationArgList();
            if (argList != null)
            {
                parameters.AddRange(argList.annotationArg().Select(arg => arg.GetText()));
            }
            return parameters;
        }

        private IAnnotation CreateAnnotation(string annotationName, IReadOnlyList<string> parameters,
            QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context)
        {
            if (_creators.TryGetValue(annotationName.ToUpperInvariant(), out var annotationClrType))
            {
                return (IAnnotation) Activator.CreateInstance(annotationClrType, qualifiedSelection, context, parameters);
            }

            return new NotRecognizedAnnotation(qualifiedSelection, context, parameters);
        }
    }
}
