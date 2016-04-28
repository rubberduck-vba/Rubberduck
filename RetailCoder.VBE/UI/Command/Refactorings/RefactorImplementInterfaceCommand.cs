using System.Linq;
using Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ImplementInterface;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorImplementInterfaceCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly ImplementInterfaceRefactoring _refactoring;

        public RefactorImplementInterfaceCommand(VBE vbe, RubberduckParserState state, IActiveCodePaneEditor editor)
            : base(vbe, editor)
        {
            _state = state;
            _refactoring = new ImplementInterfaceRefactoring(_state, Editor, new MessageBox());
        }

        public override bool CanExecute(object parameter)
        {
            return Vbe.ActiveCodePane != null && _state.Status == ParserState.Ready && _refactoring.CanExecute();
        }

        public override void Execute(object parameter)
        {
            // ReSharper disable once PossibleInvalidOperationException
            _refactoring.Refactor(Editor.GetSelection().Value);
        }
    }
}