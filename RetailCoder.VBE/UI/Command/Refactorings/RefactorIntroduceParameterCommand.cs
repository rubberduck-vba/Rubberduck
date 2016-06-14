using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.IntroduceParameter;

namespace Rubberduck.UI.Command.Refactorings
{
    public class RefactorIntroduceParameterCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;

        public RefactorIntroduceParameterCommand (VBE vbe, RubberduckParserState state)
            :base(vbe)
        {
            _state = state;
        }

        public override bool CanExecute(object parameter)
        {
            if (Vbe.ActiveCodePane == null || _state.Status != ParserState.Ready)
            {
                return false;
            }

            var selection = Vbe.ActiveCodePane.GetQualifiedSelection();
            if (!selection.HasValue)
            {
                return false;
            }

            var target = _state.AllUserDeclarations.FindVariable(selection.Value);

            var canExecute = target != null && target.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member);

            return canExecute;
        }

        public override void Execute(object parameter)
        {
            var selection = Vbe.ActiveCodePane.GetQualifiedSelection();
            if (!selection.HasValue)
            {
                return;
            }

            var refactoring = new IntroduceParameterRefactoring(Vbe, _state, new MessageBox());
            refactoring.Refactor(selection.Value);
        }
    }
}
