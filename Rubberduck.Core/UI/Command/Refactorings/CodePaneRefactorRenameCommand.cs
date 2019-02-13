using System.Runtime.InteropServices;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Rename;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class CodePaneRefactorRenameCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;
        private readonly IRefactoringPresenterFactory _factory;

        public CodePaneRefactorRenameCommand(RubberduckParserState state, IMessageBox messageBox, IRefactoringPresenterFactory factory, IRewritingManager rewritingManager, ISelectionService selectionService) 
            : base (rewritingManager, selectionService)
        {
            _state = state;
            _messageBox = messageBox;
            _factory = factory;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            if (_state.Status != ParserState.Ready)
            {
                return false;
            }

            var activeSelection = SelectionService.ActiveSelection();
            if (!activeSelection.HasValue)
            {
                return false;
            }

            var target = _state.DeclarationFinder.FindSelectedDeclaration(activeSelection.Value);

            return target != null 
                && target.IsUserDefined 
                && !_state.IsNewOrModified(target.QualifiedModuleName);
        }

        protected override void OnExecute(object parameter)
        {
            var activeSelection = SelectionService.ActiveSelection();
            if (!activeSelection.HasValue)
            {
                return;
            }

            var target = _state.DeclarationFinder.FindSelectedDeclaration(activeSelection.Value);

            if (target == null || !target.IsUserDefined)
            {
                return;
            }
            
            var refactoring = new RenameRefactoring(_factory, _messageBox, _state, _state.ProjectsProvider, RewritingManager, SelectionService);
            refactoring.Refactor(target);
        }
    }
}
