using System.Runtime.InteropServices;
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
        private readonly IGeneralConfigService _service;
        private readonly IOperatingSystem _operatingSystem;
        public SettingsCommand(IGeneralConfigService service, IOperatingSystem operatingSystem)
        {
            _service = service;
            _operatingSystem = operatingSystem;
        }

        public override void Execute(object parameter)
        {
            using (var window = new SettingsForm(_service, _operatingSystem))
            {
                window.ShowDialog();
            }
        }
    }
}
