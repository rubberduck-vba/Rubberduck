using Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ImplementInterface;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorImplementInterfaceCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly ImplementInterfaceRefactoring _refactoring;
        private QualifiedSelection? _selection;

        public RefactorImplementInterfaceCommand(VBE vbe, RubberduckParserState state, IActiveCodePaneEditor editor)
            : base(vbe, editor)
        {
            _state = state;
            _refactoring = new ImplementInterfaceRefactoring(Vbe, state, new MessageBox(), new CodePaneWrapperFactory());
        }

        public override bool CanExecute(object parameter)
        {
            _selection = Editor.GetSelection();

            return Vbe.ActiveCodePane != null && _state.Status == ParserState.Ready && _selection.HasValue && _refactoring.CanExecute(_selection.Value);
        }

        public override void Execute(object parameter)
        {
            // ReSharper disable once PossibleInvalidOperationException
            _refactoring.Refactor(_selection.Value);
        }
    }
}