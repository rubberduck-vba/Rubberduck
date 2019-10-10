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
        private readonly ISelectedDeclarationService _selectedDeclarationService;

        public CodePaneRefactorRenameCommand(
            RenameRefactoring refactoring, 
            RenameFailedNotifier renameFailedNotifier, 
            RubberduckParserState state, 
            ISelectionService selectionService,
            ISelectedDeclarationService selectedDeclarationService) 
            : base (refactoring, renameFailedNotifier, selectionService, state)
        {
            _state = state;
            _selectedDeclarationService = selectedDeclarationService;

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
            return _selectedDeclarationService.SelectedDeclaration();
        }
    }
}
