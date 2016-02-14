using System.Collections.Generic;
using System.Xml.Serialization;

namespace Rubberduck.Settings
{
    interface IGeneralSettings
    {
        DisplayLanguageSetting Language { get; set; }
        IEnumerable<Hotkey> HotkeySettings { get; set; }
    }

    [XmlType(AnonymousType = true)]
    public class GeneralSettings : IGeneralSettings
    {
        public DisplayLanguageSetting Language { get; set; }
        public IEnumerable<Hotkey> HotkeySettings { get; set; }

        public GeneralSettings()
        {
            //empty constructor needed for serialization
        }

        public GeneralSettings(DisplayLanguageSetting language, IEnumerable<Hotkey> hotkeySettings)
        {
            Language = language;
            HotkeySettings = hotkeySettings;
        }
    }
}