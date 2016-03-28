using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Annotations
{
    public interface IAnnotation
    {
        AnnotationType AnnotationType { get; }
        VBAParser.AnnotationContext Context { get; }
        AnnotationTargetType TargetType { get; }
    }
}
