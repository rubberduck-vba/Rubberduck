using System.Runtime.InteropServices;
using Rubberduck.Settings;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class TestExplorerCommand : CommandBase
    {
        private readonly IPresenter _presenter;

        public TestExplorerCommand(IPresenter presenter)
        {
            _presenter = presenter;
        }

        public override RubberduckHotkey Hotkey
        {
            get { return RubberduckHotkey.TestExplorer; }
        }

        public override void ExecuteImpl(object parameter)
        {
            _presenter.Show();
        }
    }
}
