using System;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    public abstract class RefactorCodePaneCommandBase : RefactorCommandBase
    {
        protected readonly ISelectionService SelectionService;

        protected RefactorCodePaneCommandBase(IRefactoring refactoring, IRefactoringFailureNotifier failureNotifier, ISelectionService selectionService, IParserStatusProvider parserStatusProvider)
            : base (refactoring, failureNotifier, parserStatusProvider)
        {
            SelectionService = selectionService;

            AddToCanExecuteEvaluation(SpecializedEvaluateCanExecute);
        }

        private bool SpecializedEvaluateCanExecute(object parameter)
        {
            return SelectionService.ActiveSelection().HasValue;
        }

        protected override void OnExecute(object parameter)
        {
            var activeSelection = SelectionService.ActiveSelection();
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