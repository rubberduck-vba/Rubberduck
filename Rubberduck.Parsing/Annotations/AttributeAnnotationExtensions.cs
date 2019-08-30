using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Parsing.Annotations
{
    public static class AttributeAnnotationExtensions
    {
        public static string Attribute(this ParseTreeAnnotation annotationInstance)
        {
            if (annotationInstance.Annotation is IAttributeAnnotation annotation)
            {
                return annotation.Attribute(annotationInstance.AnnotationArguments);
            }
            return null;
        }

        public static IReadOnlyList<string> AttributeValues(this ParseTreeAnnotation annotationInstance)
        {
            if (annotationInstance.Annotation is IAttributeAnnotation annotation)
            {
                return annotation.AnnotationToAttributeValues(annotationInstance.AnnotationArguments);
            }
            return null;

        }
    }
}
