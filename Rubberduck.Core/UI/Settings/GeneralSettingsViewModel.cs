using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Rubberduck.Settings;
using Rubberduck.Common;
using Rubberduck.Interaction;
using NLog;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.VbeRuntime.Settings;
using Rubberduck.Resources;

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
        private readonly IMessageBox _messageBox;
        private readonly IVbeSettings _vbeSettings;

        private bool _indenterPrompted;
        private readonly ReadOnlyCollection<Type> _experimentalFeatureTypes;

        public GeneralSettingsViewModel(Configuration config, IOperatingSystem operatingSystem, IMessageBox messageBox, IVbeSettings vbeSettings, IEnumerable<Type> experimentalFeatureTypes)
        {
            _operatingSystem = operatingSystem;
            _messageBox = messageBox;
            _vbeSettings = vbeSettings;
            _experimentalFeatureTypes = experimentalFeatureTypes.ToList().AsReadOnly();
            Languages = new ObservableCollection<DisplayLanguageSetting>(
                new[] 
            {
                new DisplayLanguageSetting("en-US"),
                new DisplayLanguageSetting("fr-CA"),
                new DisplayLanguageSetting("de-DE"),
                new DisplayLanguageSetting("cs-CZ")
            });

            LogLevels = new ObservableCollection<MinimumLogLevel>(LogLevelHelper.LogLevels.Select(l => new MinimumLogLevel(l.Ordinal, l.Name)));
            TransferSettingsToView(config.UserSettings.GeneralSettings, config.UserSettings.HotkeySettings);

            ShowLogFolderCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ShowLogFolder());
            ExportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ExportSettings());
            ImportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ImportSettings());
        }

        public List<ExperimentalFeatures> ExperimentalFeatures { get; set; }

        public ObservableCollection<DisplayLanguageSetting> Languages { get; set; } 

        private DisplayLanguageSetting _selectedLanguage;
        public DisplayLanguageSetting SelectedLanguage
        {
            get => _selectedLanguage;
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
            get => _hotkeys;
            set
            {
                if (_hotkeys != value)
                {
                    _hotkeys = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool ShouldDisplayHotkeyModificationLabel
        {
            get
            {
                return _hotkeys.Any(s => !s.IsValid);
            }
        }

        private bool _autoSaveEnabled;
        public bool AutoSaveEnabled
        {
            get => _autoSaveEnabled;
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
            get => _showSplashAtStartup;
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
            get => _checkVersionAtStartup;
            set
            {
                if (_checkVersionAtStartup != value)
                {
                    _checkVersionAtStartup = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _compileBeforeParse;
        public bool CompileBeforeParse
        {
            get => _compileBeforeParse;
            set
            {
                if (_compileBeforeParse == value)
                {
                    return;
                }

                if (value && _vbeSettings.CompileOnDemand)
                {
                    if(!SynchronizeVBESettings())
                    {
                        return;
                    }
                }

                _compileBeforeParse = value;
                OnPropertyChanged();
            }
        }

        private bool SynchronizeVBESettings()
        {
            if (!_messageBox.ConfirmYesNo(RubberduckUI.GeneralSettings_CompileBeforeParse_WarnCompileOnDemandEnabled,
                RubberduckUI.GeneralSettings_CompileBeforeParse_WarnCompileOnDemandEnabled_Caption, true))
            {
                return false;
            }

            _vbeSettings.CompileOnDemand = false;
            _vbeSettings.BackGroundCompile = false;
            return true;
        }

        private int _autoSavePeriod;
        public int AutoSavePeriod
        {
            get => _autoSavePeriod;
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
            get => _delimiter;
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
        private bool _userEditedLogLevel;

        public MinimumLogLevel SelectedLogLevel
        {
            get => _selectedLogLevel;
            set
            {
                if (!Equals(_selectedLogLevel, value))
                {
                    _userEditedLogLevel = true;
                    _selectedLogLevel = value;
                    OnPropertyChanged();
                }
            }
        }

        public CommandBase ShowLogFolderCommand { get; }

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
                CanShowSplash = ShowSplashAtStartup,
                CanCheckVersion = CheckVersionAtStartup,
                CompileBeforeParse = CompileBeforeParse,
                IsSmartIndenterPrompted = _indenterPrompted,
                IsAutoSaveEnabled = AutoSaveEnabled,
                AutoSavePeriod = AutoSavePeriod,
                UserEditedLogLevel = _userEditedLogLevel,
                MinimumLogLevel = _selectedLogLevel.Ordinal,
                EnableExperimentalFeatures = ExperimentalFeatures
            };
        }

        private void TransferSettingsToView(IGeneralSettings general, IHotkeySettings hottkey)
        {
            SelectedLanguage = Languages.First(l => l.Code == general.Language.Code);
            Hotkeys = new ObservableCollection<HotkeySetting>(hottkey.Settings);
            ShowSplashAtStartup = general.CanShowSplash;
            CheckVersionAtStartup = general.CanCheckVersion;
            CompileBeforeParse = general.CompileBeforeParse;
            _indenterPrompted = general.IsSmartIndenterPrompted;
            AutoSaveEnabled = general.IsAutoSaveEnabled;
            AutoSavePeriod = general.AutoSavePeriod;
            _userEditedLogLevel = general.UserEditedLogLevel;
            _selectedLogLevel = LogLevels.First(l => l.Ordinal == general.MinimumLogLevel);

            ExperimentalFeatures = _experimentalFeatureTypes
                .SelectMany(s => s.CustomAttributes.Where(a => a.ConstructorArguments.Any()).Select(a => (string)a.ConstructorArguments.First().Value))
                .Distinct()
                .Select(s => new ExperimentalFeatures { IsEnabled = general.EnableExperimentalFeatures.SingleOrDefault(d => d.Key == s)?.IsEnabled ?? false, Key = s })
                .ToList();
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
                general.IsSmartIndenterPrompted = true;
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