using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ImplementInterface;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorImplementInterfaceCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _msgBox;

        public RefactorImplementInterfaceCommand(RubberduckParserState state, IMessageBox msgBox, IRewritingManager rewritingManager, ISelectionService selectionService)
            : base(rewritingManager, selectionService)
        {
            _state = state;
            _msgBox = msgBox;
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

            var targetInterface = _state.DeclarationFinder.FindInterface(activeSelection.Value);
            
            var targetClass = _state.DeclarationFinder.Members(activeSelection.Value.QualifiedName)
                .SingleOrDefault(declaration => declaration.DeclarationType == DeclarationType.ClassModule);

            return targetInterface != null && targetClass != null
                && !_state.IsNewOrModified(targetInterface.QualifiedModuleName)
                && !_state.IsNewOrModified(targetClass.QualifiedModuleName);
            
        }

        protected override void OnExecute(object parameter)
        {
            var activeSelection = SelectionService.ActiveSelection();
            if (!activeSelection.HasValue)
            {
                return;
            }

            var refactoring = new ImplementInterfaceRefactoring(_state, _msgBox, RewritingManager, SelectionService);
            refactoring.Refactor(activeSelection.Value);
        }
    }
}
