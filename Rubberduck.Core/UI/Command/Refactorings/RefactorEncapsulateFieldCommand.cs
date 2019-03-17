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
    public class RefactorEncapsulateFieldCommand : RefactorCodePaneCommandBase
    {
        private readonly RubberduckParserState _state;

        public RefactorEncapsulateFieldCommand(RubberduckParserState state, Indenter indenter, IRefactoringPresenterFactory factory, IRewritingManager rewritingManager, ISelectionService selectionService)
            : base(new EncapsulateFieldRefactoring(state, indenter, factory, rewritingManager, selectionService), selectionService, state)
        {
            _state = state;

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
            var activeSelection = SelectionService.ActiveSelection();
            if (!activeSelection.HasValue)
            {
                return null;
            }

            return _state.DeclarationFinder.FindSelectedDeclaration(activeSelection.Value);
        }
    }
}
