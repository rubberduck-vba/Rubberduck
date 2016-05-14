using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections;
using Rubberduck.SmartIndenter;

namespace Rubberduck.Settings
{
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
        public Configuration LoadConfiguration()
        {
            return new Configuration
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
        }

        private List<CodeInspectionSetting> MergeImplementedInspectionsNotInConfig(List<CodeInspectionSetting> configInspections, IEnumerable<IInspection> implementedInspections)
        {
            foreach (var implementedInspection in implementedInspections)
            {
                var inspection = configInspections.SingleOrDefault(i => i.Name == implementedInspection.Name);
                if (inspection == null)
                {
                    configInspections.Add(new CodeInspectionSetting(implementedInspection));
                }
                else
                {
                    // description isn't serialized
                    inspection.Description = implementedInspection.Description;
                }
            }
            return configInspections;
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
        }

        public void SaveConfiguration(Configuration toSerialize, bool languageChanged)
        {
            SaveConfiguration(toSerialize);

            if (languageChanged)
            {
                OnLanguageChanged(EventArgs.Empty);
            }

            OnSettingsChanged(EventArgs.Empty);
        }

        public event EventHandler LanguageChanged;
        protected virtual void OnLanguageChanged(EventArgs e)
        {
            var handler = LanguageChanged;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        public event EventHandler SettingsChanged;
        protected virtual void OnSettingsChanged(EventArgs e)
        {
            var handler = SettingsChanged;
            if (handler != null)
            {
                handler(this, e);
            }
        }
    }
}
