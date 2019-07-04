using System.Runtime.InteropServices;
using Rubberduck.UI.CodeExplorer;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.UI.Command.ComCommands
{
    /// <summary>
    /// A command that displays the Code Explorer window.
    /// </summary>
    [ComVisible(false)]
    public class CodeExplorerCommand : ComCommandBase
    {
        private readonly CodeExplorerDockablePresenter _presenter;

        public CodeExplorerCommand(CodeExplorerDockablePresenter presenter, IVbeEvents vbeEvents)
            : base(vbeEvents)
        {
            _presenter = presenter;
        }

        protected override void OnExecute(object parameter)
        {
            _presenter.Show();
        }
    }
}
