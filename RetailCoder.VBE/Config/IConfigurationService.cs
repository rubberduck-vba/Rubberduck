using System;

namespace Rubberduck.Config
{
    public interface IConfigurationService<T>
    {
        T LoadConfiguration();
        void SaveConfiguration(T toSerialize);
        event EventHandler SettingsChanged;
    }
}
