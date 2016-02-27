using System.Runtime.InteropServices;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that displays the Code Explorer window.
    /// </summary>
    [ComVisible(false)]
    public class CodeExplorerCommand : CommandBase
    {
        private readonly IPresenter _presenter;

        public CodeExplorerCommand(IPresenter presenter)
        {
            _presenter = presenter;
        }

        public override void Execute(object parameter)
        {
            _presenter.Show();
        }
    }
}