using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.IntroduceField;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    public class RefactorIntroduceFieldCommand : RefactorCodePaneCommandBase
    {
        private readonly RubberduckParserState _state;

        public RefactorIntroduceFieldCommand (IntroduceFieldRefactoring refactoring, IntroduceFieldFailedNotifier introduceFieldFailedNotifier, RubberduckParserState state, ISelectionService selectionService)
            :base(refactoring, introduceFieldFailedNotifier, selectionService, state)
        {
            _state = state;

            AddToCanExecuteEvaluation(SpecializedEvaluateCanExecute);
        }

        private bool SpecializedEvaluateCanExecute(object parameter)
        {
            var target = GetTarget();

            return target != null 
                && !_state.IsNewOrModified(target.QualifiedModuleName)
                && target.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member);
        }

        private Declaration GetTarget()
        {
            var activeSelection = SelectionService.ActiveSelection();
            if (!activeSelection.HasValue)
            {
                return null;
            }

            var target = _state.DeclarationFinder
                .UserDeclarations(DeclarationType.Variable)
                .FindVariable(activeSelection.Value);
            return target;
        }
    }
}
