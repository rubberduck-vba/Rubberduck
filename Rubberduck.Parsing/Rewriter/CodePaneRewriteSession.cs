using System;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Rewriter
{
    public class CodePaneRewriteSession : RewriteSessionBase
    {
        private readonly RubberduckParserState _state;

        public CodePaneRewriteSession(RubberduckParserState state, IRewriterProvider rewriterProvider,
            Func<IRewriteSession, bool> rewritingAllowed)
            : base(rewriterProvider, rewritingAllowed)
        {
            _state = state;
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
            _state.OnParseRequested(this);
        }
    }
}