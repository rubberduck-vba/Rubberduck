using System;
using Rubberduck.UI.UnitTesting;

namespace Rubberduck.UI.Command
{
    public class TestExplorerCommand : ICommand
    {
        private readonly TestExplorerDockablePresenter _presenter;

        public TestExplorerCommand(TestExplorerDockablePresenter presenter)
        {
            _presenter = presenter;
        }

        public void Execute()
        {
            _presenter.Show();
        }

    }
}