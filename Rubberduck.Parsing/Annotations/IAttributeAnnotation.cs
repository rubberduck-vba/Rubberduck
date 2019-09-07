using System;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Annotations
{
    public interface IAttributeAnnotation : IAnnotation
    {
        bool MatchesAttributeDefinition(string attributeName, IReadOnlyList<string> attributeValues);
        string Attribute(IReadOnlyList<string> annotationValues);

        IReadOnlyList<string> AnnotationToAttributeValues(IReadOnlyList<string> annotationValues);
        IReadOnlyList<string> AttributeToAnnotationValues(IReadOnlyList<string> attributeValues);
    }
}