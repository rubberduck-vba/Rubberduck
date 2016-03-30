using System.Runtime.InteropServices;
using Rubberduck.Settings;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class TestExplorerCommand : CommandBase, IHotkeyCommand
    {
        private readonly IPresenter _presenter;

        public TestExplorerCommand(IPresenter presenter)
        {
            _presenter = presenter;
        }

        public override void Execute(object parameter)
        {
            _presenter.Show();
        }

        public RubberduckHotkey Hotkey { get {return RubberduckHotkey.TestExplorer; } }
    }
}