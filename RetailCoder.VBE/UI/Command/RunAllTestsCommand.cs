using Rubberduck.UI.UnitTesting;

namespace Rubberduck.UI.Command
{
    public class RunAllTestsCommand : ICommand
    {
        private readonly TestExplorerDockablePresenter _presenter;

        public RunAllTestsCommand(TestExplorerDockablePresenter presenter)
        {
            _presenter = presenter;
        }

        public void Execute()
        {
            _presenter.RunTests();
        }
    }
}