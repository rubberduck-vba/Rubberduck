using System;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    public abstract class RefactorCodePaneCommandBase : RefactorCommandBase
    {
        protected readonly ISelectionService SelectionService;

        protected RefactorCodePaneCommandBase(IRefactoring refactoring, ISelectionService selectionService, IParserStatusProvider parserStatusProvider)
            : base (refactoring, parserStatusProvider)
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

            Refactoring.Refactor(activeSelection.Value);
        }
    }
}