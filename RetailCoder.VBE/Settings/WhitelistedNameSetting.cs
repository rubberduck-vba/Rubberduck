using System.Xml.Serialization;

namespace Rubberduck.Settings
{
    [XmlType(AnonymousType = true)]
    public class WhitelistedNameSetting
    {
        [XmlAttribute]
        public string Name { get; set; }

        public WhitelistedNameSetting(string name)
        {
            Name = name;
        }

        public WhitelistedNameSetting() : this("*") { }
    }
}