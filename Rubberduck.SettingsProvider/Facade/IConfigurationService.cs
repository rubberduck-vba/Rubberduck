using System;

namespace Rubberduck.SettingsProvider
{
    public interface IConfigurationService<T>
    {
        T Read();
        T ReadDefaults();

        void Save(T settings);
        T Import(string fileName);
        void Export(string fileName);

        event EventHandler<ConfigurationChangedEventArgs> SettingsChanged;
    }

    public class ConfigurationChangedEventArgs : EventArgs
    {
        public bool LanguageChanged { get; }
        public bool InspectionSettingsChanged { get; }
        public bool RunInspectionsOnReparse { get; }
        public bool AutoCompleteSettingsChanged { get; }

        public ConfigurationChangedEventArgs(bool runInspections, bool languageChanged, bool inspectionSettingsChanged, bool autoCompleteSettingsChanged)
        {
            AutoCompleteSettingsChanged = autoCompleteSettingsChanged;
            RunInspectionsOnReparse = runInspections;
            LanguageChanged = languageChanged;
            InspectionSettingsChanged = inspectionSettingsChanged;
        }
    }
}