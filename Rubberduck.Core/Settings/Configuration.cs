using System.Xml.Serialization;

namespace Rubberduck.Settings
{
    [XmlType(AnonymousType = true)]
    [XmlRoot(Namespace = "", IsNullable = false)]
    public class Configuration
    {
        public UserSettings UserSettings { get; set; }

        public Configuration()
        {
            //default constructor required for serialization
        }

        public Configuration(UserSettings userSettings)
        {
            UserSettings = userSettings;
        }
    }
}
