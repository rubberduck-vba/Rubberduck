using System.Runtime.InteropServices;
using System.Windows.Input;
using Rubberduck.UI.SourceControl;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that displays the Source Control panel.
    /// </summary>
    [ComVisible(false)]
    public class ShowSourceControlPanelCommand : CommandBase
    {
        public readonly SourceControlPresenter _presenter;

        public ShowSourceControlPanelCommand(SourceControlPresenter presenter)
        {
            _presenter = presenter;
        }

        public override void Execute(object parameter)
        {
            _presenter.Show();
        }
    }
}