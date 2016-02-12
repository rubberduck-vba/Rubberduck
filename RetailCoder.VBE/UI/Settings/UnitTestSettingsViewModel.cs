using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    public class UnitTestSettingsViewModel : ViewModelBase, ISettingsViewModel
    {
        private readonly Configuration _config;

        public UnitTestSettingsViewModel(Configuration config)
        {
            _config = config;
        }

        public void UpdateConfig(Configuration config) { }
    }
}