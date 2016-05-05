using System.Runtime.InteropServices;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that displays the Source Control panel.
    /// </summary>
    [ComVisible(false)]
    public class ShowSourceControlPanelCommand : CommandBase
    {
        public readonly IPresenter _presenter;

        public ShowSourceControlPanelCommand(IPresenter presenter)
        {
            _presenter = presenter;
        }

        public override void Execute(object parameter)
        {
            _presenter.Show();
        }
    }
}