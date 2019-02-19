using System;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Rewriter
{
    public class CodePaneRewriteSession : RewriteSessionBase
    {
        private readonly IParseManager _parseManager;

        public CodePaneRewriteSession(IParseManager parseManager, IRewriterProvider rewriterProvider, ISelectionRecoverer selectionRecoverer,
            Func<IRewriteSession, bool> rewritingAllowed)
            : base(rewriterProvider, selectionRecoverer, rewritingAllowed)
        {
            _parseManager = parseManager;
        }

        public override CodeKind TargetCodeKind => CodeKind.CodePaneCode;

        protected override IExecutableModuleRewriter ModuleRewriter(QualifiedModuleName module)
        {
            return RewriterProvider.CodePaneModuleRewriter(module);
        }

        protected override bool TryRewriteInternal()
        {
            foreach (var rewriter in CheckedOutModuleRewriters.Values)
            {
                rewriter.Rewrite();
            }
            _parseManager.OnParseRequested(this);

            return true;
        }
    }
}