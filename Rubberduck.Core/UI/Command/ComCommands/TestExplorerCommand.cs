using System.Runtime.InteropServices;
using NLog;
using Rubberduck.UI.UnitTesting;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.UI.Command.ComCommands
{
    [ComVisible(false)]
    internal class TestExplorerCommand : ComCommandBase
    {
        private readonly TestExplorerDockablePresenter _presenter;

        public TestExplorerCommand(TestExplorerDockablePresenter presenter, IVbeEvents vbeEvents)
            : base(LogManager.GetCurrentClassLogger(), vbeEvents)
        {
            _presenter = presenter;
        }

        protected override void OnExecute(object parameter)
        {
            _presenter.Show();
        }
    }
}
