using System;
using Rubberduck.UI.ToDoItems;

namespace Rubberduck.UI.Command
{
    public class ToDoExplorerCommand : ICommand, IDisposable
    {
        private readonly ToDoExplorerDockablePresenter _presenter;

        public ToDoExplorerCommand(ToDoExplorerDockablePresenter presenter)
        {
            _presenter = presenter;
        }

        public void Execute()
        {
            _presenter.Show();
        }

        public void Dispose()
        {
            _presenter.Dispose();
        }
    }
}