using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.UI.Refactorings.RemoveParameters;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RemoveUnusedParameterQuickFix : QuickFixBase
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IRewritingManager _rewritingManager;
        private readonly IMessageBox _messageBox;

        public RemoveUnusedParameterQuickFix(IVBE vbe, RubberduckParserState state, IMessageBox messageBox, IRewritingManager rewritingManager)
            : base(typeof(ParameterNotUsedInspection))
        {
            _vbe = vbe;
            _state = state;
            _rewritingManager = rewritingManager;
            _messageBox = messageBox;
        }

        //The rewriteSession is optional since it is not used in this particular quickfix because it is a refactoring quickfix.
        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession = null)
        {
            using (var dialog = new RemoveParametersDialog(new RemoveParametersViewModel(_state)))
            {
                var refactoring = new RemoveParametersRefactoring(
                    _vbe,
                    new RemoveParametersPresenterFactory(_vbe, dialog, _state, _messageBox),
                    _rewritingManager);

                refactoring.QuickFix(_state, result.QualifiedSelection);
            }
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RemoveUnusedParameterQuickFix;

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => false;
    }
}