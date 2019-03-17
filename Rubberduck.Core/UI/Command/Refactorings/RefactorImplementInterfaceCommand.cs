using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ImplementInterface;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorImplementInterfaceCommand : RefactorCodePaneCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _msgBox;

        public RefactorImplementInterfaceCommand(ImplementInterfaceRefactoring refactoring, RubberduckParserState state, IMessageBox msgBox, ISelectionService selectionService)
            : base(refactoring, selectionService, state)
        {
            _state = state;
            _msgBox = msgBox;

            AddToCanExecuteEvaluation(SpecializedEvaluateCanExecute);
        }

        private bool SpecializedEvaluateCanExecute(object parameter)
        {
            var activeSelection = SelectionService.ActiveSelection();        
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
