using System;
using System.Linq;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;

namespace Rubberduck.Settings
{
    public class ConfigurationChangedEventArgs : EventArgs
    {
        public bool LanguageChanged { get; private set; }
        public bool InspectionSettingsChanged { get; private set; }
        public bool RunInspectionsOnReparse { get; private set; }

        public ConfigurationChangedEventArgs(bool runInspections, bool languageChanged, bool inspectionSettingsChanged)
        {
            RunInspectionsOnReparse = runInspections;
            LanguageChanged = languageChanged;
            InspectionSettingsChanged = inspectionSettingsChanged;
        }
    }

    public interface IGeneralConfigService : IConfigurationService<Configuration>
    {
        Configuration GetDefaultConfiguration();
    }

    public class ConfigurationLoader : IGeneralConfigService
    {
        private readonly IConfigProvider<GeneralSettings> _generalProvider;
        private readonly IConfigProvider<HotkeySettings> _hotkeyProvider;
        private readonly IConfigProvider<ToDoListSettings> _todoProvider;
        private readonly IConfigProvider<CodeInspectionSettings> _inspectionProvider;
        private readonly IConfigProvider<UnitTestSettings> _unitTestProvider;
        private readonly IConfigProvider<IndenterSettings> _indenterProvider;
        private readonly IConfigProvider<WindowSettings> _windowProvider;

        public ConfigurationLoader(IConfigProvider<GeneralSettings> generalProvider, IConfigProvider<HotkeySettings> hotkeyProvider, IConfigProvider<ToDoListSettings> todoProvider,
                                   IConfigProvider<CodeInspectionSettings> inspectionProvider, IConfigProvider<UnitTestSettings> unitTestProvider, IConfigProvider<IndenterSettings> indenterProvider, IConfigProvider<WindowSettings> windowProvider)
        {
            _generalProvider = generalProvider;
            _hotkeyProvider = hotkeyProvider;
            _todoProvider = todoProvider;
            _inspectionProvider = inspectionProvider;
            _unitTestProvider = unitTestProvider;
            _indenterProvider = indenterProvider;
            _windowProvider = windowProvider;
        }

        /// <summary>
        /// Loads the configuration from Rubberduck.config xml file.
        /// </summary>
        public virtual Configuration LoadConfiguration()
        {
            var config = new Configuration
            {
                UserSettings = new UserSettings
                (
                    _generalProvider.Create(),
                    _hotkeyProvider.Create(),
                    _todoProvider.Create(),
                    _inspectionProvider.Create(),
                    _unitTestProvider.Create(),
                    _indenterProvider.Create(),
                    _windowProvider.Create()
                )
            };            
            return config;
        }

        public Configuration GetDefaultConfiguration()
        {
            return new Configuration
            {
                UserSettings = new UserSettings
                (
                    _generalProvider.CreateDefaults(),
                    _hotkeyProvider.CreateDefaults(),
                    _todoProvider.CreateDefaults(),
                    _inspectionProvider.CreateDefaults(),
                    _unitTestProvider.CreateDefaults(),
                    _indenterProvider.CreateDefaults(),
                    _windowProvider.CreateDefaults()
                )
            };
        }
        
        public void SaveConfiguration(Configuration toSerialize)
        {
            var langChanged = _generalProvider.Create().Language.Code != toSerialize.UserSettings.GeneralSettings.Language.Code;
            var oldInspectionSettings = _inspectionProvider.Create().CodeInspections.Select(s => Tuple.Create(s.Name, s.Severity));
            var newInspectionSettings = toSerialize.UserSettings.CodeInspectionSettings.CodeInspections.Select(s => Tuple.Create(s.Name, s.Severity));
            var inspectOnReparse = toSerialize.UserSettings.CodeInspectionSettings.RunInspectionsOnSuccessfulParse;

            _generalProvider.Save(toSerialize.UserSettings.GeneralSettings);
            _hotkeyProvider.Save(toSerialize.UserSettings.HotkeySettings);
            _todoProvider.Save(toSerialize.UserSettings.ToDoListSettings);
            _inspectionProvider.Save(toSerialize.UserSettings.CodeInspectionSettings);
            _unitTestProvider.Save(toSerialize.UserSettings.UnitTestSettings);
            _indenterProvider.Save(toSerialize.UserSettings.IndenterSettings);
            _windowProvider.Save(toSerialize.UserSettings.WindowSettings);

            OnSettingsChanged(new ConfigurationChangedEventArgs(inspectOnReparse, langChanged, !oldInspectionSettings.SequenceEqual(newInspectionSettings)));
        }

        public event EventHandler<ConfigurationChangedEventArgs> SettingsChanged;
        protected virtual void OnSettingsChanged(ConfigurationChangedEventArgs e)
        {
            var handler = SettingsChanged;
            if (handler != null)
            {
                handler(this, e);
            }
        }
    }
}
