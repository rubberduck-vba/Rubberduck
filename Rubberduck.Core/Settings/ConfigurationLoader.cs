using System;
using System.Linq;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;
using Rubberduck.UnitTesting.Settings;
using Rubberduck.CodeAnalysis.Settings;

namespace Rubberduck.Settings
{

    public class ConfigurationLoader : IConfigurationService<Configuration>
    {
        private readonly IConfigurationService<GeneralSettings> _generalProvider;
        private readonly IConfigurationService<HotkeySettings> _hotkeyProvider;
        private readonly IConfigurationService<AutoCompleteSettings> _autoCompleteProvider;
        private readonly IConfigurationService<ToDoListSettings> _todoProvider;
        private readonly IConfigurationService<CodeInspectionSettings> _inspectionProvider;
        private readonly IConfigurationService<UnitTestSettings> _unitTestProvider;
        private readonly IConfigurationService<IndenterSettings> _indenterProvider;
        private readonly IConfigurationService<WindowSettings> _windowProvider;

        public ConfigurationLoader(IConfigurationService<GeneralSettings> generalProvider, 
            IConfigurationService<HotkeySettings> hotkeyProvider, 
            IConfigurationService<AutoCompleteSettings> autoCompleteProvider, 
            IConfigurationService<ToDoListSettings> todoProvider,
            IConfigurationService<CodeInspectionSettings> inspectionProvider, 
            IConfigurationService<UnitTestSettings> unitTestProvider, 
            IConfigurationService<IndenterSettings> indenterProvider, 
            IConfigurationService<WindowSettings> windowProvider)
        {
            _generalProvider = generalProvider;
            _hotkeyProvider = hotkeyProvider;
            _autoCompleteProvider = autoCompleteProvider;
            _todoProvider = todoProvider;
            _inspectionProvider = inspectionProvider;
            _unitTestProvider = unitTestProvider;
            _indenterProvider = indenterProvider;
            _windowProvider = windowProvider;
        }

        /// <summary>
        /// Loads the configuration from Rubberduck.config xml file.
        /// </summary>
        // marked virtual for Mocking
        public virtual Configuration Load()
        {
            var config = new Configuration
            {
                UserSettings = new UserSettings
                (
                    _generalProvider.Load(),
                    _hotkeyProvider.Load(),
                    _autoCompleteProvider.Load(),
                    _todoProvider.Load(),
                    _inspectionProvider.Load(),
                    _unitTestProvider.Load(),
                    _indenterProvider.Load(),
                    _windowProvider.Load()
                )
            };            
            return config;
        }

        public Configuration LoadDefaults()
        {
            return new Configuration
            {
                UserSettings = new UserSettings
                (
                    _generalProvider.LoadDefaults(),
                    _hotkeyProvider.LoadDefaults(),
                    _autoCompleteProvider.LoadDefaults(),
                    _todoProvider.LoadDefaults(),
                    _inspectionProvider.LoadDefaults(),
                    _unitTestProvider.LoadDefaults(),
                    _indenterProvider.LoadDefaults(),
                    _windowProvider.LoadDefaults()
                )
            };
        }
        
        public void Save(Configuration toSerialize)
        {
            var langChanged = _generalProvider.Load().Language.Code != toSerialize.UserSettings.GeneralSettings.Language.Code;
            var oldInspectionSettings = _inspectionProvider.Load().CodeInspections.Select(s => Tuple.Create(s.Name, s.Severity));
            var newInspectionSettings = toSerialize.UserSettings.CodeInspectionSettings.CodeInspections.Select(s => Tuple.Create(s.Name, s.Severity));
            var inspectionsChanged = !oldInspectionSettings.SequenceEqual(newInspectionSettings);
            var inspectOnReparse = toSerialize.UserSettings.CodeInspectionSettings.RunInspectionsOnSuccessfulParse;

            var oldAutoCompleteSettings = _autoCompleteProvider.Load();
            var newAutoCompleteSettings = toSerialize.UserSettings.AutoCompleteSettings;
            var autoCompletesChanged = oldAutoCompleteSettings.Equals(newAutoCompleteSettings);

            _generalProvider.Save(toSerialize.UserSettings.GeneralSettings);
            _hotkeyProvider.Save(toSerialize.UserSettings.HotkeySettings);
            _autoCompleteProvider.Save(toSerialize.UserSettings.AutoCompleteSettings);
            _todoProvider.Save(toSerialize.UserSettings.ToDoListSettings);
            _inspectionProvider.Save(toSerialize.UserSettings.CodeInspectionSettings);
            _unitTestProvider.Save(toSerialize.UserSettings.UnitTestSettings);
            _indenterProvider.Save(toSerialize.UserSettings.IndenterSettings);
            _windowProvider.Save(toSerialize.UserSettings.WindowSettings);

            OnSettingsChanged(new ConfigurationChangedEventArgs(inspectOnReparse, langChanged, inspectionsChanged, autoCompletesChanged));
        }

        public event EventHandler<ConfigurationChangedEventArgs> SettingsChanged;
        protected virtual void OnSettingsChanged(ConfigurationChangedEventArgs e)
        {
            SettingsChanged?.Invoke(this, e);
        }
    }
}
