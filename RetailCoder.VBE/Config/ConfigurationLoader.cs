﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Serialization;
using System.Runtime.InteropServices;
using System.IO;
using Rubberduck.Inspections;
using System.Reflection;
using Rubberduck.VBA.Parser.Grammar;

namespace Rubberduck.Config
{
    [ComVisible(false)]
    public static class ConfigurationLoader
    {
        private static string configFile = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Rubberduck\rubberduck.config";

        /// <summary>   Saves a Configuration to Rubberduck.config XML file via Serialization.</summary>
        public static void SaveConfiguration<T>(T toSerialize)
        {
            XmlSerializer xmlSerializer = new XmlSerializer(toSerialize.GetType());
            using (TextWriter textWriter = new StreamWriter(configFile))
            {
                xmlSerializer.Serialize(textWriter, toSerialize);
            }
        }

        /// <summary>   Loads the configuration from Rubberduck.config xml file. </summary>
        /// <remarks> If an IOException occurs returns a default configuration.</remarks>
        public static Configuration LoadConfiguration()
        {
            try
            {
                using (StreamReader reader = new StreamReader(configFile))
                {
                    var deserializer = new XmlSerializer(typeof(Configuration));
                    return (Configuration)deserializer.Deserialize(reader);
                }
            }
            catch (IOException)
            {
                return GetDefaultConfiguration();
            }
        }

        public static Configuration GetDefaultConfiguration()
        {
            var userSettings = new UserSettings(
                                    new ToDoListSettings(GetDefaultTodoMarkers()),
                                    new CodeInspectionSettings(GetDefaultCodeInspections())
                               );

            return new Configuration(userSettings);
        }

        public static ToDoMarker[] GetDefaultTodoMarkers()
        {
            var note = new ToDoMarker("NOTE:", TodoPriority.Low);
            var todo = new ToDoMarker("TODO:", TodoPriority.Normal);
            var bug = new ToDoMarker("BUG:", TodoPriority.High);

            return new ToDoMarker[] { note, todo, bug };
        }

        /// <summary>   Converts implemented code inspections into array of Config.CodeInspection objects. </summary>
        /// <returns>   An array of Config.CodeInspection. </returns>
        public static CodeInspection[] GetDefaultCodeInspections()
        {
            var configInspections = new List<CodeInspection>();
            foreach (var inspection in GetImplementedCodeInspections())
            {
                configInspections.Add(new CodeInspection(inspection));
            }

            return configInspections.ToArray();
        }

        /// <summary>   Gets all implemented code inspections via reflection </summary>
        public static IList<IInspection> GetImplementedCodeInspections()
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
        public static List<ISyntax> GetImplementedSyntax()
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
