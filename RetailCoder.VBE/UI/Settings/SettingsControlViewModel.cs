using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using Rubberduck.Settings;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Settings
{
    public class SettingsControlViewModel : ViewModelBase
    {
        private readonly IGeneralConfigService _configService;
        private readonly Configuration _config;

        public SettingsControlViewModel(IGeneralConfigService configService,
            Configuration config,
            SettingsView generalSettings,
            SettingsView todoSettings,
            SettingsView inspectionSettings,
            SettingsView unitTestSettings,
            SettingsView indenterSettings,
            SettingsViews activeView = Settings.SettingsViews.GeneralSettings)
        {
            _configService = configService;
            _config = config;

            SettingsViews = new ObservableCollection<SettingsView>
            {
                generalSettings, todoSettings, inspectionSettings, unitTestSettings, indenterSettings
            };

            SelectedSettingsView = SettingsViews.First(v => v.View == activeView);
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
            var oldLangCode = _config.UserSettings.GeneralSettings.Language.Code;

            foreach (var vm in SettingsViews.Select(v => v.Control.ViewModel))
            {
                vm.UpdateConfig(_config);
            }

            _configService.SaveConfiguration(_config, _config.UserSettings.GeneralSettings.Language.Code != oldLangCode);
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

        private ICommand _resetButtonCommand;
        public ICommand ResetButtonCommand
        {
            get
            {
                if (_resetButtonCommand != null)
                {
                    return _resetButtonCommand;
                }
                return _resetButtonCommand = new DelegateCommand(_ =>
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