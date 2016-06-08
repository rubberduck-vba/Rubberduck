using System.Collections.ObjectModel;
using System.Linq;
using Rubberduck.Settings;
using Rubberduck.Common;
using System.Windows.Input;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Settings
{
    public enum DelimiterOptions
    {
        Period = 46,
        Slash = 47
    }

    public class GeneralSettingsViewModel : ViewModelBase, ISettingsViewModel
    {
        private readonly IOperatingSystem _operatingSystem;

        public GeneralSettingsViewModel(Configuration config, IOperatingSystem operatingSystem)
        {
            _operatingSystem = operatingSystem;
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
            LogLevels = new ObservableCollection<MinimumLogLevel>(LogLevelHelper.LogLevels.Select(l => new MinimumLogLevel(l.Ordinal, l.Name)));
            SelectedLogLevel = LogLevels.First(l => l.Ordinal == config.UserSettings.GeneralSettings.MinimumLogLevel);

            _showLogFolderCommand = new DelegateCommand(_ => ShowLogFolder());
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

        public ObservableCollection<MinimumLogLevel> LogLevels { get; set; }
        private MinimumLogLevel _selectedLogLevel;
        public MinimumLogLevel SelectedLogLevel
        {
            get { return _selectedLogLevel; }
            set
            {
                if (!Equals(_selectedLogLevel, value))
                {
                    _selectedLogLevel = value;
                    OnPropertyChanged();
                }
            }
        }

        private readonly ICommand _showLogFolderCommand;
        public ICommand ShowLogFolderCommand
        {
            get { return _showLogFolderCommand; }
        }

        private void ShowLogFolder()
        {
            _operatingSystem.ShowFolder(ApplicationConstants.LOG_FOLDER_PATH);
        }

        public void UpdateConfig(Configuration config)
        {
            config.UserSettings.GeneralSettings.Language = SelectedLanguage;
            config.UserSettings.HotkeySettings.Settings = Hotkeys.ToArray();
            config.UserSettings.GeneralSettings.AutoSaveEnabled = AutoSaveEnabled;
            config.UserSettings.GeneralSettings.AutoSavePeriod = AutoSavePeriod;
            config.UserSettings.GeneralSettings.Delimiter = (char)Delimiter;
            config.UserSettings.GeneralSettings.MinimumLogLevel = SelectedLogLevel.Ordinal;
        }

        public void SetToDefaults(Configuration config)
        {
            SelectedLanguage = Languages.First(l => l.Code == config.UserSettings.GeneralSettings.Language.Code);
            Hotkeys = new ObservableCollection<HotkeySetting>(config.UserSettings.HotkeySettings.Settings);
            AutoSaveEnabled = config.UserSettings.GeneralSettings.AutoSaveEnabled;
            AutoSavePeriod = config.UserSettings.GeneralSettings.AutoSavePeriod;
            Delimiter = (DelimiterOptions)config.UserSettings.GeneralSettings.Delimiter;
            SelectedLogLevel = LogLevels.First(l => l.Ordinal == config.UserSettings.GeneralSettings.MinimumLogLevel);
        }
    }
}