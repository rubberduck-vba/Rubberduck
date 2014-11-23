using System;
using System.IO;
using System.Xml.Serialization;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Rubberduck.Config
{
    [ComVisible(false)]
    public static class ConfigurationLoader
    {
        public static Configuration LoadConfiguration()
        {
            try
            {
                string configFile = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Rubberduck\rubberduck.config";

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
            var config = new Configuration();
            var userSettings = new UserSettings();
            var todoListSettings = new ToDoListSettings();

            var note = new ToDoMarker("'NOTE:",0);
            var todo = new ToDoMarker("'TODO:", 1);
            var bug = new ToDoMarker("'BUG:", 2);

            todoListSettings.ToDoMarkers = new ToDoMarker[]{note,todo,bug};
            userSettings.ToDoListSettings = todoListSettings;
            config.UserSettings = userSettings;

            return config;
        }
    }

    [ComVisible(false)]
    [XmlTypeAttribute(AnonymousType = true)]
    [XmlRootAttribute(Namespace = "", IsNullable = false)]
    public class Configuration
    {

        public UserSettings UserSettings
        { get; set; }
    }

    [ComVisible(false)]
    [XmlTypeAttribute(AnonymousType = true)]
    public class UserSettings
    {

        public ToDoListSettings ToDoListSettings
        { get; set; }
    }
}
