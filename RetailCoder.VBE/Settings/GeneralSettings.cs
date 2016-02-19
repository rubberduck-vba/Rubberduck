using System.Xml.Serialization;

namespace Rubberduck.Settings
{
    interface IGeneralSettings
    {
        DisplayLanguageSetting Language { get; set; }
        Hotkey[] HotkeySettings { get; set; }
    }

    [XmlType(AnonymousType = true)]
    public class GeneralSettings : IGeneralSettings
    {
        public DisplayLanguageSetting Language { get; set; }

        [XmlArrayItem("Hotkey", IsNullable = false)]
        public Hotkey[] HotkeySettings { get; set; }

        public GeneralSettings()
        {
            //empty constructor needed for serialization
        }

        public GeneralSettings(DisplayLanguageSetting language, Hotkey[] hotkeySettings)
        {
            Language = language;
            HotkeySettings = hotkeySettings;
        }
    }
}