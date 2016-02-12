using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    public class UnitTestSettingsViewModel : ViewModelBase, ISettingsViewModel
    {
        public UnitTestSettingsViewModel(Configuration config)
        {
            // load from config.UserSettings.UnitTestSettings
        }

        public void UpdateConfig(Configuration config) { }
        public void SetToDefaults(Configuration config) {}
    }
}