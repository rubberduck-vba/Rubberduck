using System.Globalization;
using System.Xml.Serialization;
using Rubberduck.UI;

namespace Rubberduck.Settings
{
    [XmlType(AnonymousType = true)]
    public class DisplayLanguageSetting
    {
        [XmlAttribute]
        public string Code { get; set; }

        public DisplayLanguageSetting()
        {
            // serialization constructor
        }

        public DisplayLanguageSetting(string code)
        {
            Code = code;

            CultureInfo culture;
            try
            {
                culture = CultureInfo.GetCultureInfo(code);
                Exists = true;
            }
            catch (CultureNotFoundException)
            {
                culture = RubberduckUI.Culture;
                Exists = false;
            }

            var resource = "Language_" + Code.Substring(0, 2).ToUpper();
            Name = RubberduckUI.ResourceManager.GetString(resource, culture);
        }

        [XmlIgnore]
        public string Name { get; }

        [XmlIgnore]
        public bool Exists { get; }

        public override bool Equals(object obj)
        {
            var other = obj as DisplayLanguageSetting;
            return other != null && Code.Equals(other.Code);
        }

        public override int GetHashCode()
        {
            return Code.GetHashCode();
        }
    }
}
