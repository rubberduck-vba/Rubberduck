﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Xml.Serialization;
using Rubberduck.Inspections;

namespace Rubberduck.Config
{
    public interface IGeneralConfigService : IConfigurationService<Configuration>
    {
        CodeInspectionSetting[] GetDefaultCodeInspections();
        Configuration GetDefaultConfiguration();
        ToDoMarker[] GetDefaultTodoMarkers();
        IList<Rubberduck.Inspections.IInspection> GetImplementedCodeInspections();
    }

    public class ConfigurationLoader : XmlConfigurationServiceBase<Configuration>, IGeneralConfigService
        {

        protected override string ConfigFile
            {
            get { return Path.Combine(this.rootPath, "rubberduck.config"); }
            }

        /// <summary>   Loads the configuration from Rubberduck.config xml file. </summary>
        /// <remarks> If an IOException occurs, returns a default configuration.</remarks>
        public override Configuration LoadConfiguration()
        {
                    //deserialization can silently fail for just parts of the config, 
                    //  so we null check and return defaults if necessary.

            var config = base.LoadConfiguration();

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
                var message = ex.Message + Environment.NewLine + ex.InnerException.Message + Environment.NewLine + Environment.NewLine +
                        ConfigFile + Environment.NewLine + Environment.NewLine +
                        "Would you like to restore default configuration?" + Environment.NewLine + 
                        "Warning: All customized settings will be lost.";

            DialogResult result = MessageBox.Show(message, "Error Loading Rubberduck Configuration", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

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
            bool found;
            foreach (var implementedInspection in implementedInspections)
            {
                found = false;
                foreach (var configInspection in configInspections)
                {
                    if (implementedInspection.Name == configInspection.Name)
                    {
                        found = true;
                        break;
                    }
                }

                if (!found)
                {
                    configInspections.Add(new CodeInspectionSetting(implementedInspection));
                }
            }
            return configInspections;
        }

        public Configuration GetDefaultConfiguration()
        {
            var userSettings = new UserSettings(
                                    new ToDoListSettings(GetDefaultTodoMarkers()),
                                    new CodeInspectionSettings(GetDefaultCodeInspections())
                               );

            return new Configuration(userSettings);
        }

        public ToDoMarker[] GetDefaultTodoMarkers()
        {
            var note = new ToDoMarker("NOTE:", TodoPriority.Low);
            var todo = new ToDoMarker("TODO:", TodoPriority.Normal);
            var bug = new ToDoMarker("BUG:", TodoPriority.High);

            return new ToDoMarker[] { note, todo, bug };
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
