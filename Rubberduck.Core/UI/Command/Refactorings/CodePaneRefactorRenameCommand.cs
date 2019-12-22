using System.Runtime.InteropServices;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class CodePaneRefactorRenameCommand : RefactorCodePaneCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;

        public CodePaneRefactorRenameCommand(
            RenameRefactoring refactoring, 
            RenameFailedNotifier renameFailedNotifier, 
            RubberduckParserState state, 
            ISelectionProvider selectionProvider,
            ISelectedDeclarationProvider selectedDeclarationProvider) 
            : base (refactoring, renameFailedNotifier, selectionProvider, state)
        {
            _state = state;
            _selectedDeclarationProvider = selectedDeclarationProvider;

            AddToCanExecuteEvaluation(SpecializedEvaluateCanExecute);
        }

        private bool SpecializedEvaluateCanExecute(object parameter)
        {
            var target = GetTarget();

            return target != null
                   && target.IsUserDefined
                   && !_state.IsNewOrModified(target.QualifiedModuleName);
        }

        private Declaration GetTarget()
        {
            return _selectedDeclarationProvider.SelectedDeclaration();
        }
    }
}
