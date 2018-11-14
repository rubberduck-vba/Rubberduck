using Antlr4.Runtime;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Rewriter
{
    public interface IModuleRewriterFactory
    {
        IExecutableModuleRewriter CodePaneRewriter(QualifiedModuleName module, ITokenStream tokenStream);
        IExecutableModuleRewriter AttributesRewriter(QualifiedModuleName module, ITokenStream tokenStream);
    }
}
