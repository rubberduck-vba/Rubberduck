using System.Collections.Generic;

namespace Rubberduck.Parsing.Annotations
{
    public interface IAttributeAnnotationProvider
    {
        (IAttributeAnnotation annotation, IReadOnlyList<string> annotationValues) ModuleAttributeAnnotation(string attributeName, IReadOnlyList<string> attributeValues);
        (IAttributeAnnotation annotation, IReadOnlyList<string> annotationValues) MemberAttributeAnnotation(string attributeBaseName, IReadOnlyList<string> attributeValues);
    }
}