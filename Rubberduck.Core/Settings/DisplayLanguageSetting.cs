using System.Globalization;
using System.Xml.Serialization;
using Rubberduck.Resources;

namespace Rubberduck.Settings
{
    [XmlType(AnonymousType = true)]
    public class DisplayLanguageSetting
    {
        [XmlAttribute]
        public string Code { get; set; }

        /// <Summary>
        /// Default constructor required for XML serialization.
        /// </Summary>
        public DisplayLanguageSetting()
        {
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
            return obj is DisplayLanguageSetting other && Code.Equals(other.Code);
        }

        public override int GetHashCode()
        {
            return Code.GetHashCode();
        }
    }
}
