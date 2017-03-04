using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    public interface IAnnotationFactory
    {
        IAnnotation Create(VBAParser.AnnotationContext context, QualifiedSelection qualifiedSelection);
    }
}
