using System.Runtime.InteropServices;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that displays the To-Do explorer window.
    /// </summary>
    [ComVisible(false)]
    public class ToDoExplorerCommand : CommandBase
    {
        private readonly IPresenter _presenter;

        public ToDoExplorerCommand(IPresenter presenter)
        {
            _presenter = presenter;
        }

        public override void Execute(object parameter)
        {
            _presenter.Show();
        }
    }
}