namespace Rubberduck.UI.Command
{
    public class TestExplorerCommand : ICommand
    {
        private readonly IPresenter _presenter;

        public TestExplorerCommand(IPresenter presenter)
        {
            _presenter = presenter;
        }

        public void Execute()
        {
            _presenter.Show();
        }

    }
}