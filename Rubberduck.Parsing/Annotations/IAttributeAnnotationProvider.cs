using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Parsing.Annotations
{
    public interface IAttributeAnnotationProvider
    {
        (AnnotationType annotationType, IReadOnlyList<string> values) ModuleAttributeAnnotation(string attributeName, IReadOnlyList<string> attributeValues);
        (AnnotationType annotationType, IReadOnlyList<string> values) MemberAttributeAnnotation(string attributeName, IReadOnlyList<string> attributeValues);
    }
}