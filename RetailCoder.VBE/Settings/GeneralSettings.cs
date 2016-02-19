using System.Xml.Serialization;

namespace Rubberduck.Settings
{
    interface IGeneralSettings
    {
        DisplayLanguageSetting Language { get; set; }
        Hotkey[] HotkeySettings { get; set; }
        bool AutoSaveEnabled { get; set; }
        int AutoSavePeriod { get; set; }
    }

    [XmlType(AnonymousType = true)]
    public class GeneralSettings : IGeneralSettings
    {
        public DisplayLanguageSetting Language { get; set; }

        [XmlArrayItem("Hotkey", IsNullable = false)]
        public Hotkey[] HotkeySettings { get; set; }
        
        public bool AutoSaveEnabled { get; set; }
        public int AutoSavePeriod { get; set; }

        public GeneralSettings()
        {
            //empty constructor needed for serialization
        }

        public GeneralSettings(DisplayLanguageSetting language, Hotkey[] hotkeySettings, bool autoSaveEnabled, int autoSavePeriod)
        {
            Language = language;
            HotkeySettings = hotkeySettings;
            AutoSaveEnabled = autoSaveEnabled;
            AutoSavePeriod = autoSavePeriod;
        }
    }
}