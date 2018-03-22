using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    public interface ISettingsViewModel
    {
        void UpdateConfig(Configuration config);
        void SetToDefaults(Configuration config);
    }
}
