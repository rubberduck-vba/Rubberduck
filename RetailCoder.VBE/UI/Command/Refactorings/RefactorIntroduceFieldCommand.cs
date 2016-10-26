using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.IntroduceField;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.Refactorings
{
    public class RefactorIntroduceFieldCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public RefactorIntroduceFieldCommand (IVBE vbe, RubberduckParserState state, IMessageBox messageBox)
            :base(vbe)
        {
            _state = state;
            _messageBox = messageBox;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            var pane = Vbe.ActiveCodePane;
            {
                if (_state.Status != ParserState.Ready || pane.IsWrappingNullReference)
                {
                    return false;
                }

                var selection = pane.GetQualifiedSelection();
                if (!selection.HasValue)
                {
                    return false;
                }

                var target = _state.AllUserDeclarations.FindVariable(selection.Value);

                return target != null && target.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member);
            }
        }

        protected override void ExecuteImpl(object parameter)
        {
            var pane = Vbe.ActiveCodePane;
            {
                if (pane.IsWrappingNullReference)
                {
                    return;
                }

                var selection = pane.GetQualifiedSelection();
                if (!selection.HasValue)
                {
                    return;
                }

                var refactoring = new IntroduceFieldRefactoring(Vbe, _state, _messageBox);
                refactoring.Refactor(selection.Value);
            }
        }
    }
}
