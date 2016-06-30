using System.Runtime.InteropServices;
using Rubberduck.Settings;

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

        public override RubberduckHotkey Hotkey
        {
            get { return RubberduckHotkey.CodeExplorer; }
        }

        public override void ExecuteImpl(object parameter)
        {
            _presenter.Show();
        }
    }
}
