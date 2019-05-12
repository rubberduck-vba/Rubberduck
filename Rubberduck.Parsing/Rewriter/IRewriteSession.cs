using System.Collections.Generic;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Rewriter
{
    public interface IRewriteSession
    {
        IModuleRewriter CheckOutModuleRewriter(QualifiedModuleName module);
        IReadOnlyCollection<QualifiedModuleName> CheckedOutModules { get; }
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