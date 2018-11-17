using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Rewriter
{
    public interface IRewriterProvider
    {
        IExecutableModuleRewriter CodePaneModuleRewriter(QualifiedModuleName module);
        IExecutableModuleRewriter AttributesModuleRewriter(QualifiedModuleName module);
    }
}