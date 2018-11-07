using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Rewriter
{
    public interface IRewriteSession
    {
        IModuleRewriter CheckOutModuleRewriter(QualifiedModuleName module);
        bool TryRewrite();
        bool IsInvalidated { get; }
        void Invalidate();
        CodeKind TargetCodeKind { get; }
    }
}