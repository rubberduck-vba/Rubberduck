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

            _okButtonCommand = new DelegateCommand(_ => SaveAndCloseWindow());
            _cancelButtonCommand = new DelegateCommand(_ => CloseWindow());
            _resetButtonCommand = new DelegateCommand(_ => ResetSettings());
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

            _configService.SaveConfiguration(_config);
        }

        private void CloseWindow()
        {
            var handler = OnWindowClosed;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        private void SaveAndCloseWindow()
        {
            SaveConfig();
            CloseWindow();
        }

        private void ResetSettings()
        {
            var defaultConfig = _configService.GetDefaultConfiguration();
            foreach (var vm in SettingsViews.Select(v => v.Control.ViewModel))
            {
                vm.SetToDefaults(defaultConfig);
            }
        }

        public event EventHandler OnWindowClosed;

        private readonly ICommand _okButtonCommand;
        public ICommand OKButtonCommand
        {
            get
            {
                return _okButtonCommand;
            }
        }

        private readonly ICommand _cancelButtonCommand;
        public ICommand CancelButtonCommand
        {
            get { return _cancelButtonCommand; }
        }

        private readonly ICommand _resetButtonCommand;
        public ICommand ResetButtonCommand
        {
            get { return _resetButtonCommand; }
        }
    }
}
