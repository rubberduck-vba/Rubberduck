using System.Runtime.InteropServices;
using Rubberduck.UI.Settings;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that displays the Options dialog.
    /// </summary>
    [ComVisible(false)]
    public class SettingsCommand : CommandBase
    {
        private readonly ISettingsFormFactory _settingsFormFactory;

        public SettingsCommand(ISettingsFormFactory settingsFormFactory)
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
