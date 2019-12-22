using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorRemoveParametersCommand : RefactorCodePaneCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;

        public RefactorRemoveParametersCommand(
            RemoveParametersRefactoring refactoring, 
            RemoveParameterFailedNotifier removeParameterFailedNotifier, 
            RubberduckParserState state,
            ISelectionProvider selectionProvider,
            ISelectedDeclarationProvider selectedDeclarationProvider) 
            : base (refactoring, removeParameterFailedNotifier, selectionProvider, state)
        {
            _state = state;
            _selectedDeclarationProvider = selectedDeclarationProvider;

            AddToCanExecuteEvaluation(SpecializedEvaluateCanExecute);
        }

        private bool SpecializedEvaluateCanExecute(object parameter)
        {
            var activeSelection = SelectionProvider.ActiveSelection();
            if (!activeSelection.HasValue)
            {
                return false;
            }

            var member = GetTarget();
            if (member == null || _state.IsNewOrModified(member.QualifiedModuleName))
            {
                return false;
            }

            var parameters = _state.DeclarationFinder
                .UserDeclarations(DeclarationType.Parameter)
                .Where(item => member.Equals(item.ParentScopeDeclaration))
                .ToList();

            return member.DeclarationType == DeclarationType.PropertyLet
                   || member.DeclarationType == DeclarationType.PropertySet
                ? parameters.Count > 1
                : parameters.Any();
        }

        private Declaration GetTarget()
        {
            var selectedDeclaration = _selectedDeclarationProvider.SelectedDeclaration();
            if (!ValidDeclarationTypes.Contains(selectedDeclaration.DeclarationType))
            {
                return selectedDeclaration.DeclarationType == DeclarationType.Parameter
                    ? _selectedDeclarationProvider.SelectedMember()
                    : null;
            }

            return selectedDeclaration;
        }

        private static readonly DeclarationType[] ValidDeclarationTypes =
        {
            DeclarationType.Event,
            DeclarationType.Function,
            DeclarationType.Procedure,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };
    }
}
