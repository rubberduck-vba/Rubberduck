using System.Runtime.InteropServices;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class IndentCurrentModuleCommand : CommandBase
    {
        private readonly IIndenter _indenter;

        public IndentCurrentModuleCommand(IIndenter indenter)
        {
            _indenter = indenter;
        }

        public override void Execute(object parameter)
        {
            _indenter.IndentCurrentModule();
        }

        public RubberduckHotkey Hotkey { get {return RubberduckHotkey.IndentModule; } }
    }
}