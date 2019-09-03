using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    public interface IAnnotationFactory
    {
        IParseTreeAnnotation Create(VBAParser.AnnotationContext context, QualifiedSelection qualifiedSelection);
    }
}
