using System;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Rewriter
{
    public class CodePaneRewriteSession : RewriteSessionBase
    {
        private readonly IParseManager _parseManager;

        public CodePaneRewriteSession(IParseManager parseManager, IRewriterProvider rewriterProvider,
            Func<IRewriteSession, bool> rewritingAllowed)
            : base(rewriterProvider, rewritingAllowed)
        {
            _parseManager = parseManager;
        }


        protected override IExecutableModuleRewriter ModuleRewriter(QualifiedModuleName module)
        {
            return RewriterProvider.CodePaneModuleRewriter(module);
        }

        protected override void RewriteInternal()
        {
            foreach (var rewriter in CheckedOutModuleRewriters.Values)
            {
                rewriter.Rewrite();
            }
            _parseManager.OnParseRequested(this);
        }
    }
}