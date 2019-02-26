using System.Runtime.InteropServices;
using NLog;
using Rubberduck.UI.ToDoItems;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.UI.Command.ComCommands
{
    /// <summary>
    /// A command that displays the To-Do explorer window.
    /// </summary>
    [ComVisible(false)]
    public class ToDoExplorerCommand : ComCommandBase
    {
        private readonly ToDoExplorerDockablePresenter _presenter;

        public ToDoExplorerCommand(ToDoExplorerDockablePresenter presenter, IVBEEvents vbeEvents)
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
