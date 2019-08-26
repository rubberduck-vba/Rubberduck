using System.Collections.Generic;

namespace Rubberduck.Parsing.Annotations
{
    public interface IAttributeAnnotationProvider
    {
        (AnnotationAttribute annotationInfo, IReadOnlyList<string> values) ModuleAttributeAnnotation(string attributeName, IReadOnlyList<string> attributeValues);
        (AnnotationAttribute annotationInfo, IReadOnlyList<string> values) MemberAttributeAnnotation(string attributeBaseName, IReadOnlyList<string> attributeValues);
    }
}