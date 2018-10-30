using System;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Rewriter
{
    public class AttributesRewriteSession : RewriteSessionBase
    {
        private readonly RubberduckParserState _state;

        public AttributesRewriteSession(RubberduckParserState state, IRewriterProvider rewriterProvider,
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
            //The suspension ensures that only one parse gets executed instead of two for each rewritten module.
            var result = _state.OnSuspendParser(this, new[] {ParserState.Ready}, ExecuteAllRewriters);
            if(result != SuspensionResult.Completed)
            {
                Logger.Warn($"Rewriting attribute modules did not succeed. suspension result = {result}");
            }
        }

        private void ExecuteAllRewriters()
        {
            foreach (var rewriter in CheckedOutModuleRewriters.Values)
            {
                rewriter.Rewrite();
            }
        }
    }
}