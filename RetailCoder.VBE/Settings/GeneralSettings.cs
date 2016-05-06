using System.Xml.Serialization;

namespace Rubberduck.Settings
{
    interface IGeneralSettings
    {
        DisplayLanguageSetting Language { get; set; }
        HotkeySetting[] HotkeySettings { get; set; }
        bool AutoSaveEnabled { get; set; }
        int AutoSavePeriod { get; set; }
        char Delimiter { get; set; }
    }

    [XmlType(AnonymousType = true)]
    public class GeneralSettings : IGeneralSettings
    {
        public DisplayLanguageSetting Language { get; set; }

        [XmlArrayItem("Hotkey", IsNullable = false)]
        public HotkeySetting[] HotkeySettings { get; set; }
        
        public bool AutoSaveEnabled { get; set; }
        public int AutoSavePeriod { get; set; }

        public char Delimiter { get; set; }

        public GeneralSettings()
        {
            //empty constructor needed for serialization
        }

        public GeneralSettings(DisplayLanguageSetting language, HotkeySetting[] hotkeySettings, bool autoSaveEnabled, int autoSavePeriod, char delimiter)
        {
            Language = language;
            HotkeySettings = hotkeySettings;
            AutoSaveEnabled = autoSaveEnabled;
            AutoSavePeriod = autoSavePeriod;
            Delimiter = '.';
        }
    }
}