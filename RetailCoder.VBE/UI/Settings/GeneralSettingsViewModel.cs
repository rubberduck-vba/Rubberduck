using System.Collections.ObjectModel;
using System.Linq;
using Rubberduck.Settings;
using Rubberduck.Common;
using NLog;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Settings
{
    public enum DelimiterOptions
    {
        Period = 46,
        Slash = 47
    }

    public class GeneralSettingsViewModel : SettingsViewModelBase, ISettingsViewModel
    {
        private readonly IOperatingSystem _operatingSystem;
        private bool _indenterPrompted;

        public GeneralSettingsViewModel(Configuration config, IOperatingSystem operatingSystem)
        {
            _operatingSystem = operatingSystem;
            Languages = new ObservableCollection<DisplayLanguageSetting>(
                new[] 
            {
                new DisplayLanguageSetting("en-US"),
                new DisplayLanguageSetting("fr-CA"),
                new DisplayLanguageSetting("de-DE")
            });

            LogLevels = new ObservableCollection<MinimumLogLevel>(LogLevelHelper.LogLevels.Select(l => new MinimumLogLevel(l.Ordinal, l.Name)));
            TransferSettingsToView(config.UserSettings.GeneralSettings, config.UserSettings.HotkeySettings);

            _showLogFolderCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ShowLogFolder());
            ExportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ExportSettings());
            ImportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ImportSettings());
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

        private bool _showSplashAtStartup;
        public bool ShowSplashAtStartup
        {
            get { return _showSplashAtStartup; }
            set
            {
                if (_showSplashAtStartup != value)
                {
                    _showSplashAtStartup = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _checkVersionAtStartup;
        public bool CheckVersionAtStartup
        {
            get { return _checkVersionAtStartup; }
            set
            {
                if (_checkVersionAtStartup != value)
                {
                    _checkVersionAtStartup = value;
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

        private readonly CommandBase _showLogFolderCommand;
        public CommandBase ShowLogFolderCommand
        {
            get { return _showLogFolderCommand; }
        }

        private void ShowLogFolder()
        {
            _operatingSystem.ShowFolder(ApplicationConstants.LOG_FOLDER_PATH);
        }

        public void UpdateConfig(Configuration config)
        {
            config.UserSettings.GeneralSettings = GetCurrentGeneralSettings();
            config.UserSettings.HotkeySettings.Settings = Hotkeys.ToArray();
        }

        public void SetToDefaults(Configuration config)
        {
            TransferSettingsToView(config.UserSettings.GeneralSettings, config.UserSettings.HotkeySettings);
        }

        private Rubberduck.Settings.GeneralSettings GetCurrentGeneralSettings()
        {
            return new Rubberduck.Settings.GeneralSettings
            {
                Language = SelectedLanguage,
                ShowSplash = ShowSplashAtStartup,
                CheckVersion = CheckVersionAtStartup,
                SmartIndenterPrompted = _indenterPrompted,
                AutoSaveEnabled = AutoSaveEnabled,
                AutoSavePeriod = AutoSavePeriod,
                //Delimiter = (char)Delimiter,
                MinimumLogLevel = SelectedLogLevel.Ordinal
            };
        }

        private void TransferSettingsToView(IGeneralSettings general, IHotkeySettings hottkey)
        {
            SelectedLanguage = Languages.First(l => l.Code == general.Language.Code);
            Hotkeys = new ObservableCollection<HotkeySetting>(hottkey.Settings);
            ShowSplashAtStartup = general.ShowSplash;
            CheckVersionAtStartup = general.CheckVersion;
            _indenterPrompted = general.SmartIndenterPrompted;
            AutoSaveEnabled = general.AutoSaveEnabled;
            AutoSavePeriod = general.AutoSavePeriod;
            //Delimiter = (DelimiterOptions)general.Delimiter;
            SelectedLogLevel = LogLevels.First(l => l.Ordinal == general.MinimumLogLevel);
        }

        private void ImportSettings()
        {
            using (var dialog = new OpenFileDialog
            {
                Filter = RubberduckUI.DialogMask_XmlFilesOnly,
                Title = RubberduckUI.DialogCaption_LoadGeneralSettings
            })
            {
                dialog.ShowDialog();
                if (string.IsNullOrEmpty(dialog.FileName)) return;
                var service = new XmlPersistanceService<Rubberduck.Settings.GeneralSettings> { FilePath = dialog.FileName };
                var general = service.Load(new Rubberduck.Settings.GeneralSettings());
                var hkService = new XmlPersistanceService<HotkeySettings> { FilePath = dialog.FileName };
                var hotkey = hkService.Load(new HotkeySettings());
                //Always assume Smart Indenter registry import has been prompted if importing.
                general.SmartIndenterPrompted = true;
                TransferSettingsToView(general, hotkey);
            }
        }

        private void ExportSettings()
        {
            using (var dialog = new SaveFileDialog
            {
                Filter = RubberduckUI.DialogMask_XmlFilesOnly,
                Title = RubberduckUI.DialogCaption_SaveGeneralSettings
            })
            {
                dialog.ShowDialog();
                if (string.IsNullOrEmpty(dialog.FileName)) return;
                var service = new XmlPersistanceService<Rubberduck.Settings.GeneralSettings> { FilePath = dialog.FileName };
                service.Save(GetCurrentGeneralSettings());
                var hkService = new XmlPersistanceService<HotkeySettings> { FilePath = dialog.FileName };
                hkService.Save(new HotkeySettings { Settings = Hotkeys.ToArray() });
            }
        }
    }
}