using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Annotations
{
    public interface IAnnotationFactory
    {
        IAnnotation Create(VBAParser.AnnotationContext annotationContext, AnnotationTargetType targetType);
    }
}
