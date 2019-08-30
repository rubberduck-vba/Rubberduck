using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    public interface IAnnotationFactory
    {
        ParseTreeAnnotation Create(VBAParser.AnnotationContext context, QualifiedSelection qualifiedSelection);
    }
}
