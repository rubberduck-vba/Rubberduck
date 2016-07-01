using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using NLog;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class IndentCurrentModuleCommand : CommandBase
    {
        private readonly VBE _vbe;
        private readonly IIndenter _indenter;

        public IndentCurrentModuleCommand(VBE vbe, IIndenter indenter) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _indenter = indenter;
        }

        public override RubberduckHotkey Hotkey
        {
            get { return RubberduckHotkey.IndentModule; }
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            return _vbe.ActiveCodePane != null;
        }

        protected override void ExecuteImpl(object parameter)
        {
            _indenter.IndentCurrentModule();
        }
    }
}
