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
                _exists = true;
            }
            catch (CultureNotFoundException)
            {
                culture = RubberduckUI.Culture;
                _exists = false;
            }

            var resource = "Language_" + Code.Substring(0, 2).ToUpper();
            _name = RubberduckUI.ResourceManager.GetString(resource, culture);
        }

        private readonly string _name;
        private readonly bool _exists;

        [XmlIgnore]
        public string Name { get { return _name; } }

        [XmlIgnore]
        public bool Exists { get { return _exists; } }

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
