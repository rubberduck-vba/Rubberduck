using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ReorderParameters;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorReorderParametersCommand : RefactorCodePaneCommandBase
    {
        private readonly RubberduckParserState _state;

        public RefactorReorderParametersCommand(ReorderParametersRefactoring refactoring, ReorderParametersFailedNotifier reorderParametersFailedNotifier, RubberduckParserState state, ISelectionService selectionService) 
            : base (refactoring, reorderParametersFailedNotifier, selectionService, state)
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
            var member = _state.AllUserDeclarations.FindTarget(activeSelection.Value, ValidDeclarationTypes);
            if (member == null || _state.IsNewOrModified(member.QualifiedModuleName))
            {
                return false;
            }

            var parameters = _state.DeclarationFinder.UserDeclarations(DeclarationType.Parameter).Where(item => member.Equals(item.ParentScopeDeclaration)).ToList();
            return (member.DeclarationType == DeclarationType.PropertyLet || member.DeclarationType == DeclarationType.PropertySet)
                ? parameters.Count > 2
                : parameters.Count > 1;
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
