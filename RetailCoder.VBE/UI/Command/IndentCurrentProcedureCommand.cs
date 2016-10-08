using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class IndentCurrentProcedureCommand : CommandBase
    {
        private readonly IVBE _vbe;
        private readonly IIndenter _indenter;

        public IndentCurrentProcedureCommand(IVBE vbe, IIndenter indenter) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _indenter = indenter;
        }

        public override RubberduckHotkey Hotkey
        {
            get { return RubberduckHotkey.IndentProcedure; }
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            return _vbe.ActiveCodePane != null;
        }

        protected override void ExecuteImpl(object parameter)
        {
            _indenter.IndentCurrentProcedure();
        }
    }
}
