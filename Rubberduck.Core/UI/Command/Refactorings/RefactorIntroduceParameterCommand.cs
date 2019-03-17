using Rubberduck.Common;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.IntroduceParameter;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    public class RefactorIntroduceParameterCommand : RefactorCodePaneCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public RefactorIntroduceParameterCommand (RubberduckParserState state, IMessageBox messageBox, IRewritingManager rewritingManager, ISelectionService selectionService)
            :base(new IntroduceParameterRefactoring(state, messageBox, rewritingManager, selectionService), selectionService, state)
        {
            _state = state;
            _messageBox = messageBox;

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
