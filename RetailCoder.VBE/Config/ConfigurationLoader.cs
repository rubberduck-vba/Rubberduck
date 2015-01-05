using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Serialization;
using System.Runtime.InteropServices;
using System.IO;
using Rubberduck.Inspections;
using System.Reflection;
using Rubberduck.VBA.Parser.Grammar;
using System.Windows.Forms;

namespace Rubberduck.Config
{
    [ComVisible(false)]
    public class ConfigurationLoader : IConfigurationService
    {
        private static string configFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck", "rubberduck.config");

        /// <summary>   Saves a Configuration to Rubberduck.config XML file via Serialization.</summary>
        public void SaveConfiguration<T>(T toSerialize)
        {
            XmlSerializer xmlSerializer = new XmlSerializer(toSerialize.GetType());
            using (TextWriter textWriter = new StreamWriter(configFile))
            {
                xmlSerializer.Serialize(textWriter, toSerialize);
            }
        }

        /// <summary>   Loads the configuration from Rubberduck.config xml file. </summary>
        /// <remarks> If an IOException occurs returns a default configuration.</remarks>
        public Configuration LoadConfiguration()
        {
            try
            {
                using (StreamReader reader = new StreamReader(configFile))
                {
                    var deserializer = new XmlSerializer(typeof(Configuration));
                    var config = (Configuration)deserializer.Deserialize(reader);

                    //deserialization can silently fail for just parts of the config, 
                    //  so we null check and return defaults if necessary.
                    if (config.UserSettings.ToDoListSettings == null)
                    {
                        config.UserSettings.ToDoListSettings = new ToDoListSettings(GetDefaultTodoMarkers());
                    }

                    if (config.UserSettings.CodeInspectionSettings == null)
                    {
                        config.UserSettings.CodeInspectionSettings = new CodeInspectionSettings(GetDefaultCodeInspections());
                    }

                    //todo: check for implemented inspections that aren't in config file

                    return config;
                }
            }
            catch (IOException)
            {
                return GetDefaultConfiguration();
            }
            catch (InvalidOperationException ex)
            {
                var message = ex.Message + System.Environment.NewLine + ex.InnerException.Message + System.Environment.NewLine + System.Environment.NewLine +
                        configFile + System.Environment.NewLine + System.Environment.NewLine +
                        "Would you like to restore default configuration?" + System.Environment.NewLine + 
                        "Warning: All customized settings will be lost.";

                DialogResult result = MessageBox.Show(message, "Error Loading Rubberduck Configuration", MessageBoxButtons.YesNo,MessageBoxIcon.Exclamation);

                if (result == DialogResult.Yes)
                {
                    var config = GetDefaultConfiguration();
                    SaveConfiguration<Configuration>(config);
                    return config;
                }
                else
                {
                    throw ex;
                }
            }
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
        public CodeInspection[] GetDefaultCodeInspections()
        {
            var configInspections = new List<CodeInspection>();
            foreach (var inspection in GetImplementedCodeInspections())
            {
                configInspections.Add(new CodeInspection(inspection));
            }

            return configInspections.ToArray();
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

        /// <summary>   Gets all implemented syntax via reflection. </summary>
        public List<ISyntax> GetImplementedSyntax()
        {
            var grammar = Assembly.GetExecutingAssembly()
                                  .GetTypes()
                                  .Where(type => type.BaseType == typeof(SyntaxBase))
                                  .Select(type =>
                                  {
                                      var constructorInfo = type.GetConstructor(Type.EmptyTypes);
                                      return constructorInfo != null ? constructorInfo.Invoke(Type.EmptyTypes) : null;
                                  })
                                  .Where(syntax => syntax != null)
                                  .Cast<ISyntax>()
                                  .ToList();
            return grammar;
        }
    }
}
