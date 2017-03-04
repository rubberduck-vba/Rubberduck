using System;

namespace Rubberduck.Settings
{
    public interface IConfigurationService<T>
    {
        T LoadConfiguration();
        void SaveConfiguration(T toSerialize);
        event EventHandler<ConfigurationChangedEventArgs> SettingsChanged;
    }
}
