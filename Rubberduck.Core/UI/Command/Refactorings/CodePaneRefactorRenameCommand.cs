using System.Runtime.InteropServices;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Rename;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class CodePaneRefactorRenameCommand : RefactorCodePaneCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;
        private readonly IRefactoringPresenterFactory _factory;

        public CodePaneRefactorRenameCommand(RubberduckParserState state, IMessageBox messageBox, IRefactoringPresenterFactory factory, IRewritingManager rewritingManager, ISelectionService selectionService) 
            : base (new RenameRefactoring(factory, state, state.ProjectsProvider, rewritingManager, selectionService), selectionService, state)
        {
            _state = state;
            _messageBox = messageBox;
            _factory = factory;

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
            var activeSelection = SelectionService.ActiveSelection();
            if (!activeSelection.HasValue)
            {
                return null;
            }

            return _state.DeclarationFinder.FindSelectedDeclaration(activeSelection.Value);
        }
    }
}
