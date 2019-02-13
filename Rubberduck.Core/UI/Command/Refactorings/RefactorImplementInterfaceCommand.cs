using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ImplementInterface;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorImplementInterfaceCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IRewritingManager _rewritingManager;
        private readonly IMessageBox _msgBox;
        private readonly ISelectionService _selectionService;

        public RefactorImplementInterfaceCommand(IVBE vbe, RubberduckParserState state, IMessageBox msgBox, IRewritingManager rewritingManager, ISelectionService selectionService)
            : base(vbe)
        {
            _state = state;
            _rewritingManager = rewritingManager;
            _msgBox = msgBox;
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

            var targetInterface = _state.DeclarationFinder.FindInterface(activeSelection.Value);
            
            var targetClass = _state.DeclarationFinder.Members(activeSelection.Value.QualifiedName)
                .SingleOrDefault(declaration => declaration.DeclarationType == DeclarationType.ClassModule);

            return targetInterface != null && targetClass != null
                && !_state.IsNewOrModified(targetInterface.QualifiedModuleName)
                && !_state.IsNewOrModified(targetClass.QualifiedModuleName);
            
        }

        protected override void OnExecute(object parameter)
        {
            var activeSelection = _selectionService.ActiveSelection();
            if (!activeSelection.HasValue)
            {
                return;
            }

            var refactoring = new ImplementInterfaceRefactoring(_state, _msgBox, _rewritingManager, _selectionService);
            refactoring.Refactor(activeSelection.Value);
        }
    }
}
