using System;
using System.Collections.ObjectModel;
using System.Linq;
using NLog;
using Rubberduck.Interaction;
using Rubberduck.Resources.Settings;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Settings
{
    public class SettingsControlViewModel : ViewModelBase
    {
        private readonly IMessageBox _messageBox;
        private readonly IConfigurationService<Configuration> _configService;
        private readonly Configuration _config;

        public SettingsControlViewModel(IMessageBox messageBox,
            IConfigurationService<Configuration> configService,
            Configuration config,
            SettingsView generalSettings,
            SettingsView todoSettings,
            SettingsView inspectionSettings,
            SettingsView unitTestSettings,
            SettingsView indenterSettings,
            SettingsView autoCompleteSettings,
            SettingsView windowSettings,
            SettingsView referenceSettings,
            SettingsViews activeView = UI.Settings.SettingsViews.GeneralSettings)
        {
            _messageBox = messageBox;
            _configService = configService;
            _config = config;

            SettingsViews = new ObservableCollection<SettingsView>
            {
                generalSettings, todoSettings, inspectionSettings, unitTestSettings, indenterSettings, autoCompleteSettings, windowSettings, referenceSettings
            };

            SelectedSettingsView = SettingsViews.First(v => v.View == activeView);

            OKButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => SaveAndCloseWindow());
            CancelButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => CloseWindow());
            ResetButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ResetSettings());
        }

        private ObservableCollection<SettingsView> _settingsViews;
        public ObservableCollection<SettingsView> SettingsViews
        {
            get => _settingsViews;
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
            get => _seletedSettingsView;
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
            foreach (var vm in SettingsViews.Select(v => v.Control.ViewModel))
            {
                vm.UpdateConfig(_config);
            }

            _configService.Save(_config);
        }

        private void CloseWindow()
        {
            OnWindowClosed?.Invoke(this, EventArgs.Empty);
        }

        private void SaveAndCloseWindow()
        {
            SaveConfig();
            CloseWindow();
        }

        private void ResetSettings()
        {
            if (!_messageBox.ConfirmYesNo(SettingsUI.ConfirmResetSettings, SettingsUI.ResetSettingsButton))
            {
                return;
            }

            var defaultConfig = _configService.ReadDefaults();
            foreach (var vm in SettingsViews.Select(v => v.Control.ViewModel))
            {
                vm.SetToDefaults(defaultConfig);
            }
        }

        public event EventHandler OnWindowClosed;

        public CommandBase OKButtonCommand { get; }

        public CommandBase CancelButtonCommand { get; }

        public CommandBase ResetButtonCommand { get; }
    }
}