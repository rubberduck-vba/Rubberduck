using System.Runtime.InteropServices;
using Rubberduck.Navigation.RegexSearchReplace;
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
        {
            _presenter = presenter;
        }

        public override void Execute(object parameter)
        {
            _presenter.Show();
        }
    }

    [ComVisible(false)]
    public class RegexSearchReplaceCommand : CommandBase
    {
        private readonly RegexSearchReplacePresenter _presenter;

        public RegexSearchReplaceCommand(RegexSearchReplacePresenter presenter)
        {
            _presenter = presenter;
        }

        public override void Execute(object parameter)
        {
            _presenter.Show();
        }
    }
}