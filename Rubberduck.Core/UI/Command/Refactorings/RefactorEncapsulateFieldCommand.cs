using System.Runtime.InteropServices;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorEncapsulateFieldCommand : RefactorCodePaneCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly ISelectedDeclarationService _selectedDeclarationService;

        public RefactorEncapsulateFieldCommand(
            EncapsulateFieldRefactoring refactoring, 
            EncapsulateFieldFailedNotifier encapsulateFieldFailedNotifier, 
            RubberduckParserState state, 
            ISelectionService selectionService,
            ISelectedDeclarationService selectedDeclarationService)
            : base(refactoring, encapsulateFieldFailedNotifier, selectionService, state)
        {
            _state = state;
            _selectedDeclarationService = selectedDeclarationService;

            AddToCanExecuteEvaluation(SpecializedEvaluateCanExecute);
        }

        private bool SpecializedEvaluateCanExecute(object parameter)
        {
            var target = GetTarget();

            return target != null
                && target.DeclarationType == DeclarationType.Variable
                && !target.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member)
                && !_state.IsNewOrModified(target.QualifiedModuleName);
        }

        private Declaration GetTarget()
        {
            return _selectedDeclarationService.SelectedDeclaration();
        }
    }
}
