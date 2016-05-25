using System.Collections.ObjectModel;
using System.Linq;
using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    public enum DelimiterOptions
    {
        Period = 46,
        Slash = 47
    }

    public class GeneralSettingsViewModel : ViewModelBase, ISettingsViewModel
    {
        public GeneralSettingsViewModel(Configuration config)
        {
            Languages = new ObservableCollection<DisplayLanguageSetting>(
                new[] 
            {
                new DisplayLanguageSetting("en-US"),
                new DisplayLanguageSetting("fr-CA"),
                new DisplayLanguageSetting("de-DE"),
                new DisplayLanguageSetting("sv-SE"),
                new DisplayLanguageSetting("ja-JP")
            });

            SelectedLanguage = Languages.First(l => l.Code == config.UserSettings.GeneralSettings.Language.Code);
            Hotkeys = new ObservableCollection<HotkeySetting>(config.UserSettings.HotkeySettings.Settings);
            AutoSaveEnabled = config.UserSettings.GeneralSettings.AutoSaveEnabled;
            AutoSavePeriod = config.UserSettings.GeneralSettings.AutoSavePeriod;
            Delimiter = (DelimiterOptions)config.UserSettings.GeneralSettings.Delimiter;
            DetailedLoggingEnabled = config.UserSettings.GeneralSettings.DetailedLoggingEnabled;
        }

        public ObservableCollection<DisplayLanguageSetting> Languages { get; set; } 

        private DisplayLanguageSetting _selectedLanguage;
        public DisplayLanguageSetting SelectedLanguage
        {
            get { return _selectedLanguage; }
            set
            {
                if (!Equals(_selectedLanguage, value))
                {
                    _selectedLanguage = value;
                    OnPropertyChanged();
                }
            }
        }

        private ObservableCollection<HotkeySetting> _hotkeys;
        public ObservableCollection<HotkeySetting> Hotkeys
        {
            get { return _hotkeys; }
            set
            {
                if (_hotkeys != value)
                {
                    _hotkeys = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _autoSaveEnabled;
        public bool AutoSaveEnabled
        {
            get { return _autoSaveEnabled; }
            set
            {
                if (_autoSaveEnabled != value)
                {
                    _autoSaveEnabled = value;
                    OnPropertyChanged();
                }
            }
        }

        private int _autoSavePeriod;
        public int AutoSavePeriod
        {
            get { return _autoSavePeriod; }
            set
            {
                if (_autoSavePeriod != value)
                {
                    _autoSavePeriod = value;
                    OnPropertyChanged();
                }
            }
        }

        private DelimiterOptions _delimiter;
        public DelimiterOptions Delimiter
        {
            get { return _delimiter; }
            set
            {
                if (_delimiter != value)
                {
                    _delimiter = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _detailedLoggingEnabled;
        public bool DetailedLoggingEnabled
        {
            get { return _detailedLoggingEnabled; }
            set
            {
                if (_detailedLoggingEnabled != value)
                {
                    _detailedLoggingEnabled = value;
                    OnPropertyChanged();
                }
            }
        }

        public void UpdateConfig(Configuration config)
        {
            config.UserSettings.GeneralSettings.Language = SelectedLanguage;
            config.UserSettings.HotkeySettings.Settings = Hotkeys.ToArray();
            config.UserSettings.GeneralSettings.AutoSaveEnabled = AutoSaveEnabled;
            config.UserSettings.GeneralSettings.AutoSavePeriod = AutoSavePeriod;
            config.UserSettings.GeneralSettings.Delimiter = (char)Delimiter;
            config.UserSettings.GeneralSettings.DetailedLoggingEnabled = DetailedLoggingEnabled;
        }

        public void SetToDefaults(Configuration config)
        {
            SelectedLanguage = Languages.First(l => l.Code == config.UserSettings.GeneralSettings.Language.Code);
            Hotkeys = new ObservableCollection<HotkeySetting>(config.UserSettings.HotkeySettings.Settings);
            AutoSaveEnabled = config.UserSettings.GeneralSettings.AutoSaveEnabled;
            AutoSavePeriod = config.UserSettings.GeneralSettings.AutoSavePeriod;
            Delimiter = (DelimiterOptions)config.UserSettings.GeneralSettings.Delimiter;
            DetailedLoggingEnabled = config.UserSettings.GeneralSettings.DetailedLoggingEnabled;
        }
    }
}