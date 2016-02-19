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

        public override bool Equals(object obj)
        {
            var hotkey = obj as Hotkey;

            return hotkey != null &&
                   hotkey.Name == Name &&
                   hotkey.KeyDisplaySymbol == KeyDisplaySymbol &&
                   hotkey.IsEnabled == IsEnabled;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = (Name != null ? Name.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (KeyDisplaySymbol != null ? KeyDisplaySymbol.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ IsEnabled.GetHashCode();
                return hashCode;
            }
        }
    }
}