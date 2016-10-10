using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class IndentCurrentModuleCommand : CommandBase
    {
        private readonly IVBE _vbe;
        private readonly IIndenter _indenter;

        public IndentCurrentModuleCommand(IVBE vbe, IIndenter indenter) : base(LogManager.GetCurrentClassLogger())
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
            return !_vbe.ActiveCodePane.IsWrappingNullReference;
        }

        protected override void ExecuteImpl(object parameter)
        {
            _indenter.IndentCurrentModule();
        }
    }
}
