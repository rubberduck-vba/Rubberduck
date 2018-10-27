using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Rewriter
{
    public interface IRewriteSession
    {
        IModuleRewriter CheckOutModuleRewriter(QualifiedModuleName module);
        void Rewrite();
        bool IsInvalidated { get; }
        void Invalidate();
    }
}