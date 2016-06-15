using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections;
using Rubberduck.SmartIndenter;

namespace Rubberduck.Settings
{
    public class ConfigurationChangedEventArgs : EventArgs
    {
        public bool LanguageChanged { get; private set; }
        public bool InspectionSettingsChanged { get; private set; }

        public ConfigurationChangedEventArgs(bool languageChanged, bool inspectionSettingsChanged)
        {
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
        private readonly IGeneralConfigProvider _generalProvider;
        private readonly IHotkeyConfigProvider _hotkeyProvider;
        private readonly IToDoListConfigProvider _todoProvider;
        private readonly ICodeInspectionConfigProvider _inspectionProvider;
        private readonly IUnitTestConfigProvider _unitTestProvider;
        private readonly IIndenterConfigProvider _indenterProvider;

        private readonly IEnumerable<IInspection> _inspections;

        public ConfigurationLoader(IGeneralConfigProvider generalProvider, IHotkeyConfigProvider hotkeyProvider, IToDoListConfigProvider todoProvider,
                                   ICodeInspectionConfigProvider inspectionProvider, IUnitTestConfigProvider unitTestProvider, IIndenterConfigProvider indenterProvider,
                                   IEnumerable<IInspection> inspections)
        {
            _generalProvider = generalProvider;
            _hotkeyProvider = hotkeyProvider;
            _todoProvider = todoProvider;
            _inspectionProvider = inspectionProvider;
            _unitTestProvider = unitTestProvider;
            _indenterProvider = indenterProvider;
            _inspections = inspections;
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
                    _inspectionProvider.Create(_inspections),
                    _unitTestProvider.Create(),
                    _indenterProvider.Create()
                )
            };
            MergeImplementedInspectionsNotInConfig(config.UserSettings.CodeInspectionSettings);
            return config;
        }

        private void MergeImplementedInspectionsNotInConfig(ICodeInspectionSettings config)
        {
            foreach (var implementedInspection in _inspections)
            {
                var inspection = config.CodeInspections.SingleOrDefault(i => i.Name.Equals(implementedInspection.Name));
                if (inspection == null)
                {
                    config.CodeInspections.Add(new CodeInspectionSetting(implementedInspection));
                }
                else
                {
                    // description isn't serialized
                    inspection.Description = implementedInspection.Description;
                }
            }
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
                    _indenterProvider.CreateDefaults()
                )
            };
        }
        
        public void SaveConfiguration(Configuration toSerialize)
        {
            _generalProvider.Save(toSerialize.UserSettings.GeneralSettings);
            _hotkeyProvider.Save(toSerialize.UserSettings.HotkeySettings);
            _todoProvider.Save(toSerialize.UserSettings.ToDoListSettings);
            _inspectionProvider.Save(toSerialize.UserSettings.CodeInspectionSettings);
            _unitTestProvider.Save(toSerialize.UserSettings.UnitTestSettings);
            _indenterProvider.Save(toSerialize.UserSettings.IndenterSettings);

            var langChanged = _generalProvider.Create().Language.Code == toSerialize.UserSettings.GeneralSettings.Language.Code;
            var oldInspectionSettings = _inspectionProvider.Create(_inspections).CodeInspections.Select(s => Tuple.Create(s.Name, s.Severity));
            var newInspectionSettings = toSerialize.UserSettings.CodeInspectionSettings.CodeInspections.Select(s => Tuple.Create(s.Name, s.Severity));

            OnSettingsChanged(new ConfigurationChangedEventArgs(langChanged, oldInspectionSettings.SequenceEqual(newInspectionSettings)));
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
