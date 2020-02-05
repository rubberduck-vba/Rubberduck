using System.Collections.Generic;

namespace Rubberduck.Parsing.Annotations
{
    public static class AttributeAnnotationExtensions
    {
        public static string Attribute(this IAttributeAnnotation annotation, IParseTreeAnnotation annotationInstance)
        {
            return annotation.Attribute(annotationInstance.AnnotationArguments);
        }

        public static IReadOnlyList<string> AttributeValues(this IAttributeAnnotation annotation, IParseTreeAnnotation instance)
        {
            return annotation.AnnotationToAttributeValues(instance.AnnotationArguments);
        }
    }
}
