using System.Diagnostics;
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
        private readonly IntroduceParameterRefactoring _refactoring;

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

            return _refactoring.CanExecute(Vbe.ActiveCodePane.GetQualifiedSelection().Value);
        }

        public override void Execute(object parameter)
        {
            if (Vbe.ActiveCodePane == null)
            {
                return;
            }

            var selection = Vbe.ActiveCodePane.GetQualifiedSelection();
            _refactoring.Refactor(selection.Value);
        }
    }
}