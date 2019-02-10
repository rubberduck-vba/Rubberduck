using Rubberduck.Common;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.IntroduceField;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    public class RefactorIntroduceFieldCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IRewritingManager _rewritingManager;
        private readonly IMessageBox _messageBox;
        private readonly ISelectionService _selectionService;

        public RefactorIntroduceFieldCommand (IVBE vbe, RubberduckParserState state, IMessageBox messageBox, IRewritingManager rewritingManager, ISelectionService selectionService)
            :base(vbe)
        {
            _state = state;
            _rewritingManager = rewritingManager;
            _messageBox = messageBox;
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

            var target = _state.AllUserDeclarations.FindVariable(activeSelection.Value);

            return target != null 
                && !_state.IsNewOrModified(target.QualifiedModuleName)
                && target.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member);
        }

        protected override void OnExecute(object parameter)
        {
            var activeSelection = _selectionService.ActiveSelection();
            if (!activeSelection.HasValue)
            {
                return;
            }

            var refactoring = new IntroduceFieldRefactoring(_state, _messageBox, _rewritingManager, _selectionService);
            refactoring.Refactor(activeSelection.Value);
        }
    }
}
