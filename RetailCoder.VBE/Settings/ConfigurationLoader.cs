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
        //        /// <summary>
        ///// Defines the root path where all Rubberduck Configuration files are stored.
        ///// </summary>
        //protected readonly string rootPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck");

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
        /// <remarks>
        /// Returns default configuration when an IOException is caught.
        /// </remarks>
        public Configuration LoadConfiguration()
        {
            //deserialization can silently fail for just parts of the config, 
            //so we null-check and return defaults if necessary.
            return new Configuration
            {
                UserSettings = new UserSettings
                {
                    GeneralSettings = _generalProvider.Create(),
                    HotkeySettings = _hotkeyProvider.Create(),
                    ToDoListSettings = _todoProvider.Create(),
                    CodeInspectionSettings = _inspectionProvider.Create(_inspections),
                    UnitTestSettings = _unitTestProvider.Create(),
                    IndenterSettings = _indenterProvider.Create()
                }
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
                {
                    GeneralSettings = _generalProvider.CreateDefaults(),
                    HotkeySettings = _hotkeyProvider.CreateDefaults(),
                    ToDoListSettings = _todoProvider.CreateDefaults(),
                    CodeInspectionSettings = _inspectionProvider.CreateDefaults(),
                    UnitTestSettings = _unitTestProvider.CreateDefaults(),
                    IndenterSettings = _indenterProvider.CreateDefaults()
                }
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
