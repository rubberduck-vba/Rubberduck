using System.Linq;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveCloserToUsage;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    public class RefactorMoveCloserToUsageCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IRewritingManager _rewritingManager;
        private readonly IMessageBox _msgbox;
        private readonly ISelectionService _selectionService;

        public RefactorMoveCloserToUsageCommand(IVBE vbe, RubberduckParserState state, IMessageBox msgbox, IRewritingManager rewritingManager, ISelectionService selectionService)
            :base(vbe)
        {
            _state = state;
            _rewritingManager = rewritingManager;
            _msgbox = msgbox;
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

            var target = _state.DeclarationFinder.FindSelectedDeclaration(activeSelection.Value);
            return target != null
                   && !_state.IsNewOrModified(target.QualifiedModuleName)
                   && (target.DeclarationType == DeclarationType.Variable
                       || target.DeclarationType == DeclarationType.Constant)
                   && target.References.Any();
        }

        protected override void OnExecute(object parameter)
        {
            var activeSelection = _selectionService.ActiveSelection();

            if (activeSelection.HasValue)
            {
                var refactoring = new MoveCloserToUsageRefactoring(_state, _msgbox, _rewritingManager, _selectionService);
                refactoring.Refactor(activeSelection.Value);
            }
        }
    }
}
