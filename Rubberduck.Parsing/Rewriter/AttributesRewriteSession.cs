using System;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Rewriter
{
    public class AttributesRewriteSession : RewriteSessionBase
    {
        private readonly IParseManager _parseManager;

        public AttributesRewriteSession(IParseManager parseManager, IRewriterProvider rewriterProvider,
            Func<IRewriteSession, bool> rewritingAllowed)
            : base(rewriterProvider, rewritingAllowed)
        {
            _parseManager = parseManager;
        }

        public override CodeKind TargetCodeKind => CodeKind.AttributesCode;

        protected override IExecutableModuleRewriter ModuleRewriter(QualifiedModuleName module)
        {
            return RewriterProvider.AttributesModuleRewriter(module);
        }

        protected override bool TryRewriteInternal()
        {
            //The suspension ensures that only one parse gets executed instead of two for each rewritten module.
            var result = _parseManager.OnSuspendParser(this, new[] {ParserState.Ready}, ExecuteAllRewriters);
            if(result != SuspensionResult.Completed)
            {
                Logger.Warn($"Rewriting attribute modules did not succeed. suspension result = {result}");
                return false;
            }

            return true;
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