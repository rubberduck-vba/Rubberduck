using System.Linq;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveCloserToUsage;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.Refactorings
{
    public class RefactorMoveCloserToUsageCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IRewritingManager _rewritingManager;
        private readonly IMessageBox _msgbox;

        public RefactorMoveCloserToUsageCommand(IVBE vbe, RubberduckParserState state, IMessageBox msgbox, IRewritingManager rewritingManager)
            :base(vbe)
        {
            _state = state;
            _rewritingManager = rewritingManager;
            _msgbox = msgbox;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            using (var activePane = Vbe.ActiveCodePane)
            {
                if (activePane == null || activePane .IsWrappingNullReference || _state.Status != ParserState.Ready)
                {
                    return false;
                }

                var target = _state.FindSelectedDeclaration(activePane);
                return target != null
                       && !_state.IsNewOrModified(target.QualifiedModuleName)
                       && (target.DeclarationType == DeclarationType.Variable ||
                           target.DeclarationType == DeclarationType.Constant)
                       && target.References.Any();
            }
        }

        protected override void OnExecute(object parameter)
        {
            var selection = Vbe.GetActiveSelection();

            if (selection.HasValue)
            {
                var refactoring = new MoveCloserToUsageRefactoring(Vbe, _state, _msgbox, _rewritingManager);
                refactoring.Refactor(selection.Value);
            }
        }
    }
}
