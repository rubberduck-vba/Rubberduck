using System.Runtime.InteropServices;
using NLog;
using Rubberduck.UI.UnitTesting;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    internal class TestExplorerCommand : CommandBase
    {
        private readonly TestExplorerDockablePresenter _presenter;

        public TestExplorerCommand(TestExplorerDockablePresenter presenter)
            : base(LogManager.GetCurrentClassLogger())
        {
            _presenter = presenter;
        }

        protected override void OnExecute(object parameter)
        {
            _presenter.Show();
        }
    }
}
