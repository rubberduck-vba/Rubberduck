using System.Collections.Generic;

namespace Rubberduck.Parsing.Annotations
{
    public interface IAttributeAnnotation : IAnnotation
    {
        string Attribute { get; }
        IReadOnlyList<string> AttributeValues { get; }
    }
}