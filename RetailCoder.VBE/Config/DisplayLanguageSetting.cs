using System.Globalization;
using System.Xml.Serialization;
using Rubberduck.UI;

namespace Rubberduck.Config
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
            _name = RubberduckUI.ResourceManager.GetString("Language_" + Code.Substring(0, 2).ToUpper());
        }

        private readonly string _name;

        [XmlIgnore]
        public string Name { get { return _name; } }

        public override bool Equals(object obj)
        {
            var other = (DisplayLanguageSetting) obj;
            return Code.Equals(other.Code);
        }

        public override int GetHashCode()
        {
            return Code.GetHashCode();
        }
    }
}