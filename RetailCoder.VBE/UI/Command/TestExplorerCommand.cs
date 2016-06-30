using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Settings;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class TestExplorerCommand : CommandBase
    {
        private readonly IPresenter _presenter;

        public TestExplorerCommand(IPresenter presenter) : base(LogManager.GetCurrentClassLogger())
        {
            _presenter = presenter;
        }

        public override RubberduckHotkey Hotkey
        {
            get { return RubberduckHotkey.TestExplorer; }
        }

        protected override void ExecuteImpl(object parameter)
        {
            _presenter.Show();
        }
    }
}
