using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    public interface ISettingsViewModel<out TSettings> : ISettingsViewModel
        where TSettings : class, new() 
    { }

    public interface ISettingsViewModel
    {
        void UpdateConfig(Configuration config);
        void SetToDefaults(Configuration config);
    }
}
