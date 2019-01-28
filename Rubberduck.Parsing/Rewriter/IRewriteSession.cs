using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Rewriter
{
    public interface IRewriteSession
    {
        IModuleRewriter CheckOutModuleRewriter(QualifiedModuleName module);
        bool TryRewrite();
        RewriteSessionState Status { get; set; }
        CodeKind TargetCodeKind { get; }
    }

    public enum RewriteSessionState
    {
        Valid,
        RewriteApplied,
        OtherSessionsRewriteApplied,
        StaleParseTree
    }
}