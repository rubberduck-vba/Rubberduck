using System.Collections.Generic;

namespace Rubberduck.Parsing.Annotations
{
    public interface IAttributeAnnotationProvider
    {
        (AnnotationType annotationType, IReadOnlyList<string> values) ModuleAttributeAnnotation(string attributeName, IReadOnlyList<string> attributeValues);
        (AnnotationType annotationType, IReadOnlyList<string> values) MemberAttributeAnnotation(string attributeBaseName, IReadOnlyList<string> attributeValues);
    }
}