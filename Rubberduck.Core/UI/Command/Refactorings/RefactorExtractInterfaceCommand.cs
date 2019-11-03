using System.Runtime.InteropServices;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorExtractInterfaceCommand : RefactorCodePaneCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly ExtractInterfaceRefactoring _extractInterfaceRefactoring;

        public RefactorExtractInterfaceCommand(
            ExtractInterfaceRefactoring refactoring, 
            ExtractInterfaceFailedNotifier extractInterfaceFailedNotifier, 
            RubberduckParserState state, 
            ISelectionProvider selectionProvider)
            :base(refactoring, extractInterfaceFailedNotifier, selectionProvider, state)
        {
            _state = state;
            _extractInterfaceRefactoring = refactoring;

            AddToCanExecuteEvaluation(SpecializedEvaluateCanExecute);
        }

        private bool SpecializedEvaluateCanExecute(object parameter)
        {
            var activeSelection = SelectionProvider.ActiveSelection();
            if (!activeSelection.HasValue)
            {
                return false;
            }
            return _extractInterfaceRefactoring.CanExecute(_state, activeSelection.Value.QualifiedName);
        }
    }
}
