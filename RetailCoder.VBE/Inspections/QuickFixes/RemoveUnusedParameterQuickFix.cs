using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.UI;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Inspections.QuickFixes
{
    public class RemoveUnusedParameterQuickFix : QuickFixBase
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public RemoveUnusedParameterQuickFix(ParserRuleContext context, QualifiedSelection selection, 
            IVBE vbe, RubberduckParserState state, IMessageBox messageBox)
            : base(context, selection, InspectionsUI.RemoveUnusedParameterQuickFix)
        {
            _vbe = vbe;
            _state = state;
            _messageBox = messageBox;
        }

        public override void Fix()
        {
            using (var dialog = new RemoveParametersDialog())
            {
                var refactoring = new RemoveParametersRefactoring(_vbe,
                    new RemoveParametersPresenterFactory(_vbe, dialog, _state, _messageBox));

                refactoring.QuickFix(_state, Selection);
            }
        }
    }
}