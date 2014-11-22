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

    [System.Runtime.InteropServices.ComVisible(false)]
    [XmlTypeAttribute(AnonymousType = true)]
    [XmlRootAttribute(Namespace = "", IsNullable = false)]
    public class Configuration
    {

        public UserSettings UserSettings
        { get; set; }
    }

    [System.Runtime.InteropServices.ComVisible(false)]
    [XmlTypeAttribute(AnonymousType = true)]
    public class UserSettings
    {

        public ToDoListSettings ToDoListSettings
        { get; set; }
    }
}
