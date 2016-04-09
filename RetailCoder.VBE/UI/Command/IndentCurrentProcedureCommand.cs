using System.Runtime.InteropServices;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class IndentCurrentProcedureCommand : CommandBase
    {
        private readonly IIndenter _indenter;

        public IndentCurrentProcedureCommand(IIndenter indenter)
        {
            _indenter = indenter;
        }

        public override void Execute(object parameter)
        {
            _indenter.IndentCurrentProcedure();
        }

        public RubberduckHotkey Hotkey { get {return RubberduckHotkey.IndentProcedure; } }
    }
}