using Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorExtractInterfaceCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public RefactorExtractInterfaceCommand(VBE vbe, RubberduckParserState state, IActiveCodePaneEditor editor, IMessageBox messageBox)
            : base (vbe, editor)
        {
            _state = state;
            _messageBox = messageBox;
        }

        public override void Execute(object parameter)
        {
            if (Vbe.ActiveCodePane == null)
            {
                return;
            }

            using (var view = new ExtractInterfaceDialog())
            {
                var factory = new ExtractInterfacePresenterFactory(_state, Editor, view);
                var refactoring = new ExtractInterfaceRefactoring(_state, _messageBox, factory, Editor);
                refactoring.Refactor();
            }
        }
    }
}