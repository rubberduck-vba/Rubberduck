using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Settings;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that displays the Code Explorer window.
    /// </summary>
    [ComVisible(false)]
    public class CodeExplorerCommand : CommandBase
    {
        private readonly IDockablePresenter _presenter;

        public CodeExplorerCommand(IDockablePresenter presenter)
            : base(LogManager.GetCurrentClassLogger())
        {
            _presenter = presenter;
        }

        public override HotkeySetting DefaultHotkey => new HotkeySetting(typeof(CodeExplorerCommand))
        {
            IsEnabled = true,
            HasCtrlModifier = true,
            Key1 = "R"
        };

        protected override void OnExecute(object parameter)
        {
            _presenter.Show();
        }
    }
}
