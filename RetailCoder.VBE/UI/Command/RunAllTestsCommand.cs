using Rubberduck.UI.UnitTesting;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that runs all Rubberduck unit tests in the VBE.
    /// </summary>
    public class RunAllTestsCommand : CommandBase
    {
        private readonly TestExplorerDockablePresenter _presenter;

        public RunAllTestsCommand(TestExplorerDockablePresenter presenter)
        {
            _presenter = presenter;
        }

        public override void Execute(object parameter)
        {
            _presenter.RunTests();
        }
    }
}