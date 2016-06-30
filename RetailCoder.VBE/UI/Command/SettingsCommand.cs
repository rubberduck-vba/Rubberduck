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
        private readonly IGeneralConfigService _service;
        private readonly IOperatingSystem _operatingSystem;
        public SettingsCommand(IGeneralConfigService service, IOperatingSystem operatingSystem) : base(LogManager.GetCurrentClassLogger())
        {
            _service = service;
            _operatingSystem = operatingSystem;
        }

        protected override void ExecuteImpl(object parameter)
        {
            using (var window = new SettingsForm(_service, _operatingSystem))
            {
                window.ShowDialog();
            }
        }
    }
}
