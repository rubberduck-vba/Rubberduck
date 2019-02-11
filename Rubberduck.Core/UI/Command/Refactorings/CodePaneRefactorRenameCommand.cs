using System.Runtime.InteropServices;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Refactorings.Rename;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class CodePaneRefactorRenameCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IRewritingManager _rewritingManager;
        private readonly IMessageBox _messageBox;
        private readonly IRefactoringPresenterFactory _factory;
        private readonly ISelectionService _selectionService;

        public CodePaneRefactorRenameCommand(IVBE vbe, RubberduckParserState state, IMessageBox messageBox, IRefactoringPresenterFactory factory, IRewritingManager rewritingManager, ISelectionService selectionService) 
            : base (vbe)
        {
            _state = state;
            _rewritingManager = rewritingManager;
            _messageBox = messageBox;
            _factory = factory;
            _selectionService = selectionService;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            if (_state.Status != ParserState.Ready)
            {
                return false;
            }

            var activeSelection = _selectionService.ActiveSelection();
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
            var activeSelection = _selectionService.ActiveSelection();
            if (!activeSelection.HasValue)
            {
                return;
            }

            var target = _state.DeclarationFinder.FindSelectedDeclaration(activeSelection.Value);

            if (target == null || !target.IsUserDefined)
            {
                return;
            }
            
            var refactoring = new RenameRefactoring(_factory, _messageBox, _state, _state.ProjectsProvider, _rewritingManager, _selectionService);
            refactoring.Refactor(target);
        }
    }
}
