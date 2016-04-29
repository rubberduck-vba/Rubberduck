using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.IntroduceParameter;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Command.Refactorings
{
    public class RefactorIntroduceParameterCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IntroduceParameterRefactoring _refactoring;
        private QualifiedSelection _qualifiedSelection;

        public RefactorIntroduceParameterCommand (VBE vbe, RubberduckParserState state)
            :base(vbe)
        {
            _state = state;
            _refactoring = new IntroduceParameterRefactoring(Vbe, _state, new MessageBox());
        }

        public override bool CanExecute(object parameter)
        {
            if (Vbe.ActiveCodePane == null || _state.Status != ParserState.Ready)
            {
                return false;
            }

            var qualifiedSelection = Vbe.ActiveCodePane.GetQualifiedSelection();

            if (qualifiedSelection == null)
            {
                return false;
            }

            _qualifiedSelection = qualifiedSelection.Value;

            return _refactoring.CanExecute(_qualifiedSelection);
        }

        public override void Execute(object parameter)
        {
            _refactoring.Refactor(_qualifiedSelection);
        }
    }
}