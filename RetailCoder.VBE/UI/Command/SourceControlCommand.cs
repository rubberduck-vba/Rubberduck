using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Settings;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that displays the Source Control panel.
    /// </summary>
    [ComVisible(false)]
    public class SourceControlCommand : CommandBase
    {
        private readonly IDockablePresenter _presenter;

        public SourceControlCommand(IDockablePresenter presenter) : base(LogManager.GetCurrentClassLogger())
        {
            _presenter = presenter;
        }

        protected override void OnExecute(object parameter)
        {
            _presenter.Show();
        }

        public override RubberduckHotkey Hotkey => RubberduckHotkey.SourceControl;
    }
}
