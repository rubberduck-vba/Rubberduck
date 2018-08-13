using Antlr4.Runtime;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Rewriter
{
    public interface IModuleRewriterFactory
    {
        IModuleRewriter CodePaneRewriter(QualifiedModuleName module, ITokenStream tokenStream);
        IModuleRewriter AttributesRewriter(QualifiedModuleName module, ITokenStream tokenStream);
    }
}
