using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Common;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ReorderParameters;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorReorderParametersCommand : RefactorCodePaneCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;
        private readonly IRefactoringPresenterFactory _factory;

        public RefactorReorderParametersCommand(RubberduckParserState state, IRefactoringPresenterFactory factory, IMessageBox messageBox, IRewritingManager rewritingManager, ISelectionService selectionService) 
            : base (new ReorderParametersRefactoring(state, factory, rewritingManager, selectionService), selectionService, state)
        {
            _state = state;
            _messageBox = messageBox;

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
