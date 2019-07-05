using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Common;
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

        public RefactorRemoveParametersCommand(RemoveParametersRefactoring refactoring, RemoveParameterFailedNotifier removeParameterFailedNotifier, RubberduckParserState state, ISelectionService selectionService) 
            : base (refactoring, removeParameterFailedNotifier, selectionService, state)
        {
            _state = state;

            AddToCanExecuteEvaluation(SpecializedEvaluateCanExecute);
        }

        private bool SpecializedEvaluateCanExecute(object parameter)
        {
            var activeSelection = SelectionService.ActiveSelection();
            if (!activeSelection.HasValue)
            {
                return false;
            }

            var member = _state.DeclarationFinder.AllUserDeclarations.FindTarget(activeSelection.Value, ValidDeclarationTypes);
            if (member == null)
            {
                return false;
            }

            if (_state.IsNewOrModified(member.QualifiedModuleName))
            {
                return false;
            }

            var parameters = _state.DeclarationFinder.UserDeclarations(DeclarationType.Parameter)
                .Where(item => member.Equals(item.ParentScopeDeclaration))
                .ToList();
            return member.DeclarationType == DeclarationType.PropertyLet
                   || member.DeclarationType == DeclarationType.PropertySet
                ? parameters.Count > 1
                : parameters.Any();
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
