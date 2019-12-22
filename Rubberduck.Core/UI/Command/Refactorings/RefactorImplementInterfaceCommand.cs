using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ImplementInterface;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorImplementInterfaceCommand : RefactorCodePaneCommandBase
    {
        private readonly RubberduckParserState _state;

        public RefactorImplementInterfaceCommand(
            ImplementInterfaceRefactoring refactoring, 
            ImplementInterfaceFailedNotifier implementInterfaceFailedNotifier, 
            RubberduckParserState state,
            ISelectionProvider selectionProvider)
            : base(refactoring, implementInterfaceFailedNotifier, selectionProvider, state)
        {
            _state = state;

            AddToCanExecuteEvaluation(SpecializedEvaluateCanExecute);
        }

        private bool SpecializedEvaluateCanExecute(object parameter)
        {
            var activeSelection = SelectionProvider.ActiveSelection();        
            if (!activeSelection.HasValue)
            {
                return false;
            }

            var targetInterface = _state.DeclarationFinder.FindInterface(activeSelection.Value);
            
            var targetClass = _state.DeclarationFinder.Members(activeSelection.Value.QualifiedName)
                .SingleOrDefault(declaration => declaration.DeclarationType == DeclarationType.ClassModule);

            return targetInterface != null && targetClass != null
                && !_state.IsNewOrModified(targetInterface.QualifiedModuleName)
                && !_state.IsNewOrModified(targetClass.QualifiedModuleName); 
        }
    }
}
