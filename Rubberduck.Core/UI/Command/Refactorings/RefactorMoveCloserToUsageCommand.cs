using System.Linq;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveCloserToUsage;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    public class RefactorMoveCloserToUsageCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _msgbox;

        public RefactorMoveCloserToUsageCommand(RubberduckParserState state, IMessageBox msgbox, IRewritingManager rewritingManager, ISelectionService selectionService)
            :base(rewritingManager, selectionService)
        {
            _state = state;
            _msgbox = msgbox;
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

            var target = _state.DeclarationFinder.FindSelectedDeclaration(activeSelection.Value);
            return target != null
                   && !_state.IsNewOrModified(target.QualifiedModuleName)
                   && (target.DeclarationType == DeclarationType.Variable
                       || target.DeclarationType == DeclarationType.Constant)
                   && target.References.Any();
        }

        protected override void OnExecute(object parameter)
        {
            var activeSelection = SelectionService.ActiveSelection();
            if (activeSelection.HasValue)
            {
                var refactoring = new MoveCloserToUsageRefactoring(_state, _msgbox, RewritingManager, SelectionService);
                refactoring.Refactor(activeSelection.Value);
            }
        }
    }
}
