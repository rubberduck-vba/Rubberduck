using System;

namespace Rubberduck.Settings
{
    public interface IConfigurationService<T>
    {
        T LoadConfiguration();
        void SaveConfiguration(T toSerialize);
        void SaveConfiguration(T toSerialize, bool languageChanged);
        event EventHandler LanguageChanged;
        event EventHandler SettingsChanged;
    }
}
