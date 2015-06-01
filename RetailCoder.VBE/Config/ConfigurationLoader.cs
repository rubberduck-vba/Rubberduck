using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Rubberduck.Inspections;
using Rubberduck.UI;

namespace Rubberduck.Config
{
    public interface IGeneralConfigService : IConfigurationService<Configuration>
    {
        CodeInspectionSetting[] GetDefaultCodeInspections();
        Configuration GetDefaultConfiguration();
        ToDoMarker[] GetDefaultTodoMarkers();
        IList<IInspection> GetImplementedCodeInspections();
    }

    public class ConfigurationLoader : XmlConfigurationServiceBase<Configuration>, IGeneralConfigService
    {
        public event EventHandler SettingsChanged;

        protected override string ConfigFile
        {
            get { return Path.Combine(rootPath, "rubberduck.config"); }
        }

        /// <summary>
        /// Loads the configuration from Rubberduck.config xml file.
        /// </summary>
        /// <remarks>
        /// Returns default configuration when an IOException is caught.
        /// </remarks>
        public override Configuration LoadConfiguration()
        {
            //deserialization can silently fail for just parts of the config, 
            //so we null-check and return defaults if necessary.

            var config = base.LoadConfiguration();

            if (config.UserSettings.LanguageSetting == null)
            {
                config.UserSettings.LanguageSetting = new DisplayLanguageSetting("en-US");
            }

            if (config.UserSettings.ToDoListSettings == null)
            {
                config.UserSettings.ToDoListSettings = new ToDoListSettings(GetDefaultTodoMarkers());
            }

            if (config.UserSettings.CodeInspectionSettings == null)
            {
                config.UserSettings.CodeInspectionSettings = new CodeInspectionSettings(GetDefaultCodeInspections());
            }

            var implementedInspections = GetImplementedCodeInspections();
            var configInspections = config.UserSettings.CodeInspectionSettings.CodeInspections.ToList();
            
            configInspections = MergeImplementedInspectionsNotInConfig(configInspections, implementedInspections);
            config.UserSettings.CodeInspectionSettings.CodeInspections = configInspections.ToArray();

            return config;
        }

        protected override Configuration HandleIOException(IOException ex)
        {
            return GetDefaultConfiguration();
        }

        protected override Configuration HandleInvalidOperationException(InvalidOperationException ex)
        {
            var message = string.Format(RubberduckUI.PromptLoadDefaultConfig, ex.Message, ex.InnerException.Message, ConfigFile);
            var result = MessageBox.Show(message, RubberduckUI.LoadConfigError, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

            if (result == DialogResult.Yes)
            {
                var config = GetDefaultConfiguration();
                SaveConfiguration(config);
                return config;
            }

            throw ex;
        }

        private List<CodeInspectionSetting> MergeImplementedInspectionsNotInConfig(List<CodeInspectionSetting> configInspections, IList<IInspection> implementedInspections)
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
            var userSettings = new UserSettings(
                                    new DisplayLanguageSetting("en-US"), 
                                    new ToDoListSettings(GetDefaultTodoMarkers()),
                                    new CodeInspectionSettings(GetDefaultCodeInspections())
                               );

            return new Configuration(userSettings);
        }

        public ToDoMarker[] GetDefaultTodoMarkers()
        {
            var note = new ToDoMarker(RubberduckUI.ToDoMarkerNote, TodoPriority.Low);
            var todo = new ToDoMarker(RubberduckUI.ToDoMarkerToDo, TodoPriority.Normal);
            var bug = new ToDoMarker(RubberduckUI.ToDoMarkerBug, TodoPriority.High);

            return new[] { note, todo, bug };
        }

        /// <summary>   Converts implemented code inspections into array of Config.CodeInspection objects. </summary>
        /// <returns>   An array of Config.CodeInspection. </returns>
        public CodeInspectionSetting[] GetDefaultCodeInspections()
        {
            return GetImplementedCodeInspections()
                    .Select(x => new CodeInspectionSetting(x))
                    .ToArray();
        }

        /// <summary>   Gets all implemented code inspections via reflection </summary>
        public IList<IInspection> GetImplementedCodeInspections()
        {
            var inspections = Assembly.GetExecutingAssembly()
                                  .GetTypes()
                                  .Where(type => type.GetInterfaces().Contains(typeof(IInspection)))
                                  .Select(type =>
                                  {
                                      var constructor = type.GetConstructor(Type.EmptyTypes);
                                      return constructor != null ? constructor.Invoke(Type.EmptyTypes) : null;
                                  })
                                 .Where(inspection => inspection != null)
                                  .Cast<IInspection>()
                                  .ToList();

            return inspections;
        }
    }
}
