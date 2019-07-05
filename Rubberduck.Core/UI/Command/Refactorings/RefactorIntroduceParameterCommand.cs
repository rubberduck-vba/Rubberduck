using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.IntroduceParameter;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    public class RefactorIntroduceParameterCommand : RefactorCodePaneCommandBase
    {
        private readonly RubberduckParserState _state;

        public RefactorIntroduceParameterCommand (IntroduceParameterRefactoring refactoring, IntroduceParameterFailedNotifier introduceParameterFailedNotifier, RubberduckParserState state, ISelectionService selectionService)
            :base(refactoring, introduceParameterFailedNotifier, selectionService, state)
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
