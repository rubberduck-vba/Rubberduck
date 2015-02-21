using System;
using System.IO;
using System.Xml.Serialization;
using System.Runtime.InteropServices;

namespace Rubberduck.Config
{
    [XmlTypeAttribute(AnonymousType = true)]
    [XmlRootAttribute(Namespace = "", IsNullable = false)]
    public class Configuration
    {
        public UserSettings UserSettings { get; set; }

        public Configuration()
        {
            //default constructor required for serialization
        }

        public Configuration(UserSettings userSettings)
        {
            this.UserSettings = userSettings;
        }
    }
}
