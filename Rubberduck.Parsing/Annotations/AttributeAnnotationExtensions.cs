using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
