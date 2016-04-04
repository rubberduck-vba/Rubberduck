using Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.Settings;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorExtractMethodCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;

        public RefactorExtractMethodCommand(VBE vbe, RubberduckParserState state, IActiveCodePaneEditor editor)
            : base (vbe, editor)
        {
            _state = state;
        }

        public override void Execute(object parameter)
        {
            var factory = new ExtractMethodPresenterFactory(Editor, _state.AllDeclarations);
            var refactoring = new ExtractMethodRefactoring(factory, Editor);
            refactoring.InvalidSelection += HandleInvalidSelection;
            refactoring.Refactor();
        }

        public RubberduckHotkey Hotkey { get {return RubberduckHotkey.RefactorExtractMethod; } }
    }
}