using System.Linq;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveCloserToUsage;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    public class RefactorMoveCloserToUsageCommand : RefactorCodePaneCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public RefactorMoveCloserToUsageCommand(MoveCloserToUsageRefactoring refactoring, RubberduckParserState state, IMessageBox messageBox, ISelectionService selectionService)
            :base(refactoring, selectionService, state)
        {
            _state = state;
            _messageBox = messageBox;

            AddToCanExecuteEvaluation(SpecializedEvaluateCanExecute);
        }

        private bool SpecializedEvaluateCanExecute(object parameter)
        {
            var target = GetTarget();

            return target != null
                   && !_state.IsNewOrModified(target.QualifiedModuleName)
                   && (target.DeclarationType == DeclarationType.Variable
                       || target.DeclarationType == DeclarationType.Constant)
                   && target.References.Any();
        }

        private Declaration GetTarget()
        {
            var activeSelection = SelectionService.ActiveSelection();
            if (!activeSelection.HasValue)
            {
                return null;
            }

            var target = _state.DeclarationFinder.FindSelectedDeclaration(activeSelection.Value);
            return target;
        }
    }
}
