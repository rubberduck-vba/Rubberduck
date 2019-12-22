using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    public abstract class RefactorCodePaneCommandBase : RefactorCommandBase
    {
        protected readonly ISelectionProvider SelectionProvider;

        protected RefactorCodePaneCommandBase(
            IRefactoring refactoring, 
            IRefactoringFailureNotifier failureNotifier, 
            ISelectionProvider selectionProvider, 
            IParserStatusProvider parserStatusProvider)
            : base (refactoring, failureNotifier, parserStatusProvider)
        {
            SelectionProvider = selectionProvider;

            AddToCanExecuteEvaluation(SpecializedEvaluateCanExecute);
        }

        private bool SpecializedEvaluateCanExecute(object parameter)
        {
            return SelectionProvider.ActiveSelection().HasValue;
        }

        protected override void OnExecute(object parameter)
        {
            var activeSelection = SelectionProvider.ActiveSelection();
            if (!activeSelection.HasValue)
            {
                return;
            }

            try
            {
                Refactoring.Refactor(activeSelection.Value);
            }
            catch (RefactoringAbortedException)
            {}
            catch (RefactoringException exception)
            {
                FailureNotifier.Notify(exception);
            }
        }
    }
}