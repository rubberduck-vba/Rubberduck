using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Inspections.Abstract;
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
        private readonly IMessageBox _messageBox;

        public RemoveUnusedParameterQuickFix(IVBE vbe, RubberduckParserState state, IMessageBox messageBox)
            : base(typeof(ParameterNotUsedInspection))
        {
            _vbe = vbe;
            _state = state;
            _messageBox = messageBox;
        }

        public override void Fix(IInspectionResult result)
        {
            using (var dialog = new RemoveParametersDialog(new RemoveParametersViewModel(_state)))
            {
                var refactoring = new RemoveParametersRefactoring(_vbe,
                    new RemoveParametersPresenterFactory(_vbe, dialog, _state, _messageBox));

                refactoring.QuickFix(_state, result.QualifiedSelection);
            }
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RemoveUnusedParameterQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}