using System.Diagnostics;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.IntroduceField;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.Command.Refactorings
{
    public class RefactorIntroduceFieldCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IntroduceFieldRefactoring _refactoring;
        private QualifiedSelection _qualifiedSelection;

        public RefactorIntroduceFieldCommand (VBE vbe, RubberduckParserState state)
            :base(vbe)
        {
            _state = state;
            _refactoring = new IntroduceFieldRefactoring(Vbe, _state, new MessageBox());
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