using System.Runtime.InteropServices;
using Rubberduck.Settings;
using Rubberduck.UI.Settings;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that displays the Options dialog.
    /// </summary>
    [ComVisible(false)]
    public class SettingsCommand : CommandBase
    {
        private readonly IGeneralConfigService _service;
        public SettingsCommand(IGeneralConfigService service)
        {
            _service = service;
        }

        public override void Execute(object parameter)
        {
            using (var window = new SettingsForm(_service))
            {
                window.ShowDialog();
            }
        }
    }
}
