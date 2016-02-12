using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    public class UnitTestSettingsViewModel : ViewModelBase
    {
        private readonly IGeneralConfigService _configService;
        private readonly Configuration _config;

        public UnitTestSettingsViewModel(IGeneralConfigService configService)
        {
            _configService = configService;
            _config = configService.LoadConfiguration();
        }
    }
}