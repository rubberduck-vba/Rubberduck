using System.Runtime.InteropServices;

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

        public override void Execute(object parameter)
        {
            _presenter.Show();
        }
    }
}