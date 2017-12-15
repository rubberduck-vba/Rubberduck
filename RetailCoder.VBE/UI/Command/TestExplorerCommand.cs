using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Settings;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class TestExplorerCommand : CommandBase
    {
        private readonly IDockablePresenter _presenter;

        public TestExplorerCommand(IDockablePresenter presenter)
            : base(LogManager.GetCurrentClassLogger())
        {
            _presenter = presenter;
        }

        public override RubberduckHotkey Hotkey => RubberduckHotkey.TestExplorer;

        protected override void OnExecute(object parameter)
        {
            _presenter.Show();
        }
    }
}
