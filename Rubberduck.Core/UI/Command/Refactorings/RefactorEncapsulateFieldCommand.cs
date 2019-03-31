using System.Runtime.InteropServices;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorEncapsulateFieldCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly Indenter _indenter;
        private readonly IRefactoringPresenterFactory _factory;

        public RefactorEncapsulateFieldCommand(RubberduckParserState state, Indenter indenter, IRefactoringPresenterFactory factory, IRewritingManager rewritingManager, ISelectionService selectionService)
            : base(rewritingManager, selectionService)
        {
            _state = state;
            _indenter = indenter;
            _factory = factory;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            //This should come first because it does not require COM access.
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
                && target.DeclarationType == DeclarationType.Variable
                && !target.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member)
                && !_state.IsNewOrModified(target.QualifiedModuleName);
        }

        protected override void OnExecute(object parameter)
        {
            if(!SelectionService.ActiveSelection().HasValue)
            {
                return;
            }

            var refactoring = new EncapsulateFieldRefactoring(_state, _indenter, _factory, RewritingManager, SelectionService);
            refactoring.Refactor();
        }
    }
}
