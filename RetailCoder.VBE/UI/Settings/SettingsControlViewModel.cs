using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using Rubberduck.Settings;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Settings
{
    public class SettingsView
    {
        public string Label { get; set; }
        public string Instructions
        { 
            get
            {
                return RubberduckUI.ResourceManager.GetString("SettingsInstructions_" + View);
            } 
        }
        public ISettingsView Control { get; set; }
        public SettingsViews View { get; set; }
    }

    public class SettingsControlViewModel : ViewModelBase
    {
        private readonly IGeneralConfigService _configService;
        private readonly Configuration _config;

        public SettingsControlViewModel(IGeneralConfigService configService, SettingsViews view = Settings.SettingsViews.GeneralSettings)
        {
            _configService = configService;
            _config = configService.LoadConfiguration();

            SettingsViews = new ObservableCollection<SettingsView>
            {
                new SettingsView
                {
                    Label = RubberduckUI.SettingsCaption_GeneralSettings,
                    Control = new GeneralSettings(new GeneralSettingsViewModel(_config)),
                    View = Settings.SettingsViews.GeneralSettings
                },
                new SettingsView
                {
                    Label = RubberduckUI.SettingsCaption_TodoSettings,
                    Control = new TodoSettings(new TodoSettingsViewModel(_config)),
                    View = Settings.SettingsViews.TodoSettings
                },
                new SettingsView
                {
                    Label = RubberduckUI.SettingsCaption_CodeInspections,
                    Control = new InspectionSettings(new InspectionSettingsViewModel(_config)),
                    View = Settings.SettingsViews.InspectionSettings
                },
                new SettingsView
                {
                    Label = RubberduckUI.SettingsCaption_UnitTestSettings,
                    Control = new UnitTestSettings(new UnitTestSettingsViewModel(_config)),
                    View = Settings.SettingsViews.UnitTestSettings
                },
                /*new SettingsView
                {
                    Label = RubberduckUI.SettingsCaption_IndenterSettings,
                    Control = new GeneralSettings(),
                    View = Settings.SettingsViews.IndenterSettings
                }*/
            };

            SelectedSettingsView = SettingsViews.First(v => v.View == view);
        }

        private ObservableCollection<SettingsView> _settingsViews;
        public ObservableCollection<SettingsView> SettingsViews
        {
            get
            {
                return _settingsViews;
            }
            set
            {
                if (_settingsViews != value)
                {
                    _settingsViews = value;
                    OnPropertyChanged();
                }
            }
        }

        private SettingsView _seletedSettingsView;
        public SettingsView SelectedSettingsView
        {
            get { return _seletedSettingsView; }
            set
            {
                if (_seletedSettingsView != value)
                {
                    _seletedSettingsView = value;
                    OnPropertyChanged();
                }
            }
        }

        private void SaveConfig()
        {
            var oldLangCode = _config.UserSettings.LanguageSetting.Code;

            foreach (var vm in SettingsViews.Select(v => v.Control.ViewModel))
            {
                vm.UpdateConfig(_config);
            }

            _configService.SaveConfiguration(_config, _config.UserSettings.LanguageSetting.Code != oldLangCode);
        }

        public event EventHandler OnOKButtonClicked;
        public event EventHandler OnCancelButtonClicked;

        #region Commands

        private ICommand _okButtonCommand;
        public ICommand OKButtonCommand
        {
            get
            {
                if (_okButtonCommand != null)
                {
                    return _okButtonCommand;
                }
                return _okButtonCommand = new DelegateCommand(_ =>
                {
                    SaveConfig();

                    var handler = OnOKButtonClicked;
                    if (handler != null)
                    {
                        handler(this, EventArgs.Empty);
                    }
                });
            }
        }

        private ICommand _cancelButtonCommand;
        public ICommand CancelButtonCommand
        {
            get
            {
                if (_cancelButtonCommand != null)
                {
                    return _cancelButtonCommand;
                }
                return _cancelButtonCommand = new DelegateCommand(_ =>
                {
                    var handler = OnCancelButtonClicked;
                    if (handler != null)
                    {
                        handler(this, EventArgs.Empty);
                    }
                });
            }
        }

        private ICommand _refreshButtonCommand;
        public ICommand RefreshButtonCommand
        {
            get
            {
                if (_refreshButtonCommand != null)
                {
                    return _refreshButtonCommand;
                }
                return _refreshButtonCommand = new DelegateCommand(_ =>
                {
                    var defaultConfig = _configService.GetDefaultConfiguration();
                    foreach (var vm in SettingsViews.Select(v => v.Control.ViewModel))
                    {
                        vm.SetToDefaults(defaultConfig);
                    }
                });
            }
        }

        #endregion
    }
}