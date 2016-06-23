using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class IndentCurrentModuleCommand : CommandBase
    {
        private readonly VBE _vbe;
        private readonly IIndenter _indenter;

        public IndentCurrentModuleCommand(VBE vbe, IIndenter indenter)
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
            _indenter.IndentCurrentModule();
        }

        public RubberduckHotkey Hotkey { get { return RubberduckHotkey.IndentModule; } }
    }
}
