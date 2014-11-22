using System;
using System.IO;
using System.Xml.Serialization;

namespace Rubberduck.Config
{
    [System.Runtime.InteropServices.ComVisible(false)]
    public static class ConfigurationLoader
    {
        public static Configuration LoadConfiguration()
        {
            string configFile = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\Rubberduck\rubberduck.config";

            using (StreamReader reader = new StreamReader(configFile))
            {
                var deserializer = new XmlSerializer(typeof(Configuration));
                return (Configuration)deserializer.Deserialize(reader);
            }
        }
    }

    //todo: remove noise; use default setters/getters
    [System.Runtime.InteropServices.ComVisible(false)]
    [XmlTypeAttribute(AnonymousType = true)]
    [XmlRootAttribute(Namespace = "", IsNullable = false)]
    public  class Configuration
    {
        private UserSettings userSettingsField;

        public UserSettings UserSettings
        {
            get
            {
                return this.userSettingsField;
            }
            set
            {
                this.userSettingsField = value;
            }
        }
    }

    [System.Runtime.InteropServices.ComVisible(false)]
    [XmlTypeAttribute(AnonymousType = true)]
    public class UserSettings
    {

        private ToDoListSettings toDoListSettingsField;

        public ToDoListSettings ToDoListSettings
        {
            get
            {
                return this.toDoListSettingsField;
            }
            set
            {
                this.toDoListSettingsField = value;
            }
        }
    }
}
