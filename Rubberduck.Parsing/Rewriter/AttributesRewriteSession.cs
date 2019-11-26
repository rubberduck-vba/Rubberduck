using System;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Rewriter
{
    public class AttributesRewriteSession : RewriteSessionBase
    {
        private readonly IParseManager _parseManager;

        public AttributesRewriteSession(IParseManager parseManager, IRewriterProvider rewriterProvider, ISelectionRecoverer selectionRecoverer,
            Func<IRewriteSession, bool> rewritingAllowed)
            : base(rewriterProvider, selectionRecoverer, rewritingAllowed)
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
            GuaranteeReparseAfterRewrite();
            PrimeActiveCodePaneRecovery();
            //Attribute rewrites close the affected code panes, so we have to recover the open state.
            PrimeOpenStateRecovery();

            var result = _parseManager.OnSuspendParser(this, new[] {ParserState.Ready, ParserState.ResolvedDeclarations}, ExecuteAllRewriters);
            if(result.Outcome != SuspensionOutcome.Completed)
            {
                Logger.Warn($"Rewriting attribute modules did not succeed. Suspension result = {result}");
                if (result.EncounteredException != null)
                {
                    Logger.Warn(result.EncounteredException);
                }
                return false;
            }

            return true;
        }

        private void GuaranteeReparseAfterRewrite()
        {
            _parseManager.StateChanged += ReparseOnSuspension;
        }

        private void ReparseOnSuspension(object requestor, ParserStateEventArgs e)
        {
            if (e.State != ParserState.Busy)
            {
                return;
            }

            _parseManager.StateChanged -= ReparseOnSuspension;
            _parseManager.OnParseRequested(this);
        }

        private void PrimeActiveCodePaneRecovery()
        {
            SelectionRecoverer.SaveActiveCodePane();
            SelectionRecoverer.RecoverActiveCodePaneOnNextParse();
        }

        private void PrimeOpenStateRecovery()
        {
            SelectionRecoverer.SaveOpenState(CheckedOutModules);
            SelectionRecoverer.RecoverOpenStateOnNextParse();
        }

        private void ExecuteAllRewriters()
        {
            foreach (var module in CheckedOutModuleRewriters.Keys)
            {
                //We have to mark the modules explicitly as modified because attributes only changes do not alter the code pane code.
                _parseManager.MarkAsModified(module);
                CheckedOutModuleRewriters[module].Rewrite();
            }
        }
    }
}