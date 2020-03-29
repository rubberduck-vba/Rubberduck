using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveFolder;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    public class CodePaneRefactorMoveContainingFolderCommand : RefactorCodePaneCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;

        public CodePaneRefactorMoveContainingFolderCommand(
            MoveContainingFolderRefactoring refactoring,
            MoveContainingFolderRefactoringFailedNotifier failureNotifier,
            ISelectionProvider selectionProvider,
            RubberduckParserState state,
            ISelectedDeclarationProvider selectedDeclarationProvider)
            : base(refactoring, failureNotifier, selectionProvider, state)
        {
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _state = state;

            AddToCanExecuteEvaluation(SpecializedEvaluateCanExecute);
        }

        private bool SpecializedEvaluateCanExecute(object parameter)
        {
            var target = GetTarget();

            return target is ModuleDeclaration
                   && !_state.IsNewOrModified(target.QualifiedModuleName);
        }

        private Declaration GetTarget()
        {
            return _selectedDeclarationProvider.SelectedModule();
        }
    }
}