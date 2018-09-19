using System.Runtime.InteropServices;
using NLog;
using Rubberduck.UI.ToDoItems;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that displays the To-Do explorer window.
    /// </summary>
    [ComVisible(false)]
    public class ToDoExplorerCommand : CommandBase
    {
        private readonly ToDoExplorerDockablePresenter _presenter;

        public ToDoExplorerCommand(ToDoExplorerDockablePresenter presenter)
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
