using System.Xml.Serialization;
using Rubberduck.UI;

namespace Rubberduck.Settings
{
    public class Hotkey
    {
        public string Name { get; set; }
        public string KeyDisplaySymbol { get; set; }
        public bool IsEnabled { get; set; }

        [XmlIgnore]
        public string Prompt
        {
            get { return RubberduckUI.ResourceManager.GetString(Name + "Hotkey_Description"); } 
        }
    }
}