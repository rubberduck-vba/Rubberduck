using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Settings;
using Rubberduck.UI.Settings;
using Rubberduck.Common;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that displays the Options dialog.
    /// </summary>
    [ComVisible(false)]
    public class SettingsCommand : CommandBase
    {
        private readonly ISettingsFormFactory _settingsFormFactory;

        public SettingsCommand(ISettingsFormFactory settingsFormFactory) : base(LogManager.GetCurrentClassLogger())
        {
            _settingsFormFactory = settingsFormFactory;
        }

        protected override void OnExecute(object parameter)
        {
            using (var window = _settingsFormFactory.Create())
            {
                window.ShowDialog();
                _settingsFormFactory.Release(window);
            }
        }
    }
}
