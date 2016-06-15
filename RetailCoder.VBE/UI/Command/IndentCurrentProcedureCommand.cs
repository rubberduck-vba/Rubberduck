using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class IndentCurrentProcedureCommand : CommandBase
    {
        private readonly VBE _vbe;
        private readonly IIndenter _indenter;

        public IndentCurrentProcedureCommand(VBE vbe, IIndenter indenter)
        {
            _vbe = vbe;
            _indenter = indenter;
        }

        public override bool CanExecute(object parameter)
        {
            return _vbe.ActiveCodePane != null;
        }

        public override void Execute(object parameter)
        {
            _indenter.IndentCurrentProcedure();
        }

        public RubberduckHotkey Hotkey { get {return RubberduckHotkey.IndentProcedure; } }
    }
}
