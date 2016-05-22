using System.Xml.Serialization;

namespace Rubberduck.Settings
{
    public interface IGeneralSettings
    {
        DisplayLanguageSetting Language { get; set; }
        bool AutoSaveEnabled { get; set; }
        int AutoSavePeriod { get; set; }
        char Delimiter { get; set; }
        bool DetailedLoggingEnabled { get; set; }
    }

    [XmlType(AnonymousType = true)]
    public class GeneralSettings : IGeneralSettings
    {
        public DisplayLanguageSetting Language { get; set; }
        public bool AutoSaveEnabled { get; set; }
        public int AutoSavePeriod { get; set; }
        public char Delimiter { get; set; }
        public bool DetailedLoggingEnabled { get; set; }

        public GeneralSettings()
        {
            Language = new DisplayLanguageSetting("en-US");
            AutoSaveEnabled = false;
            AutoSavePeriod = 10;
            Delimiter = '.';
            DetailedLoggingEnabled = false;
        }

        public GeneralSettings(
            DisplayLanguageSetting language, 
            bool autoSaveEnabled, 
            int autoSavePeriod, 
            bool detailedLoggingEnabled)
        {
            Language = language;
            AutoSaveEnabled = autoSaveEnabled;
            AutoSavePeriod = autoSavePeriod;
            Delimiter = '.';
            DetailedLoggingEnabled = detailedLoggingEnabled;
        }
    }
}