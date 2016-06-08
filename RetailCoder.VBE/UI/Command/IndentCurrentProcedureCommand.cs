using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class IndentCurrentProcedureCommand : CommandBase
    {
        private readonly VBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IIndenter _indenter;

        public IndentCurrentProcedureCommand(VBE vbe, RubberduckParserState state, IIndenter indenter)
        {
            _vbe = vbe;
            _state = state;
            _indenter = indenter;
        }

        public override bool CanExecute(object parameter)
        {
            return _state.FindSelectedDeclaration(_vbe.ActiveCodePane, true) != null;
        }

        public override void Execute(object parameter)
        {
            _indenter.IndentCurrentProcedure();
        }

        public RubberduckHotkey Hotkey { get {return RubberduckHotkey.IndentProcedure; } }
    }
}
