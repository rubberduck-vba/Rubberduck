using System.Xml.Serialization;

namespace Rubberduck.CodeAnalysis.Settings
{
    [XmlType(AnonymousType = true)]
    public class WhitelistedIdentifierSetting
    {
        [XmlAttribute]
        public string Identifier { get; set; }

        public WhitelistedIdentifierSetting(string identifier)
        {
            Identifier = identifier;
        }

        public WhitelistedIdentifierSetting() : this("*") { }
    }
}