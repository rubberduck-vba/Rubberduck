using NLog;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.Command;
using Rubberduck.Resources.Settings;

namespace Rubberduck.UI.Settings
{
    public sealed class WindowSettingsViewModel : SettingsViewModelBase<Rubberduck.Settings.WindowSettings>, ISettingsViewModel<Rubberduck.Settings.WindowSettings>
    {
        public WindowSettingsViewModel(Configuration config, IConfigurationService<Rubberduck.Settings.WindowSettings> service) 
            : base(service)
        {
            CodeExplorerVisibleOnStartup = config.UserSettings.WindowSettings.CodeExplorerVisibleOnStartup;
            CodeInspectionsVisibleOnStartup = config.UserSettings.WindowSettings.CodeInspectionsVisibleOnStartup;
            TestExplorerVisibleOnStartup = config.UserSettings.WindowSettings.TestExplorerVisibleOnStartup;
            TodoExplorerVisibleOnStartup = config.UserSettings.WindowSettings.TodoExplorerVisibleOnStartup;

            ExportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ =>
                ExportSettings(new Rubberduck.Settings.WindowSettings()
                {
                    CodeExplorerVisibleOnStartup = CodeExplorerVisibleOnStartup,
                    CodeInspectionsVisibleOnStartup = CodeInspectionsVisibleOnStartup,
                    TestExplorerVisibleOnStartup = TestExplorerVisibleOnStartup,
                    TodoExplorerVisibleOnStartup = TodoExplorerVisibleOnStartup
                }));
            ImportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ImportSettings());
        }

        #region Properties

        private bool _codeExplorerVisibleOnStartup;
        public bool CodeExplorerVisibleOnStartup
        {
            get => _codeExplorerVisibleOnStartup;
            set
            {
                if (_codeExplorerVisibleOnStartup != value)
                {
                    _codeExplorerVisibleOnStartup = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _codeInspectionsVisibleOnStartup;
        public bool CodeInspectionsVisibleOnStartup
        {
            get => _codeInspectionsVisibleOnStartup;
            set
            {
                if (_codeInspectionsVisibleOnStartup != value)
                {
                    _codeInspectionsVisibleOnStartup = value;
                    OnPropertyChanged();
                }
            }
        }
        
        private bool _testExplorerVisibleOnStartup;
        public bool TestExplorerVisibleOnStartup
        {
            get => _testExplorerVisibleOnStartup;
            set
            {
                if (_testExplorerVisibleOnStartup != value)
                {
                    _testExplorerVisibleOnStartup = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _todoExplorerVisibleOnStartup;
        public bool TodoExplorerVisibleOnStartup
        {
            get => _todoExplorerVisibleOnStartup;
            set
            {
                if (_todoExplorerVisibleOnStartup != value)
                {
                    _todoExplorerVisibleOnStartup = value;
                    OnPropertyChanged();
                }
            }
        }

        #endregion

        public void UpdateConfig(Configuration config)
        {
            config.UserSettings.WindowSettings.CodeExplorerVisibleOnStartup = CodeExplorerVisibleOnStartup;
            config.UserSettings.WindowSettings.CodeInspectionsVisibleOnStartup = CodeInspectionsVisibleOnStartup;
            config.UserSettings.WindowSettings.TestExplorerVisibleOnStartup = TestExplorerVisibleOnStartup;
            config.UserSettings.WindowSettings.TodoExplorerVisibleOnStartup = TodoExplorerVisibleOnStartup;
        }

        public void SetToDefaults(Configuration config)
        {
            TransferSettingsToView(config.UserSettings.WindowSettings);
        }

        protected override string DialogLoadTitle => SettingsUI.DialogCaption_LoadWindowSettings;
        protected override string DialogSaveTitle => SettingsUI.DialogCaption_SaveWindowSettings;

        protected override void TransferSettingsToView(Rubberduck.Settings.WindowSettings toLoad)
        {
            CodeExplorerVisibleOnStartup = toLoad.CodeExplorerVisibleOnStartup;
            CodeInspectionsVisibleOnStartup = toLoad.CodeInspectionsVisibleOnStartup;
            TestExplorerVisibleOnStartup = toLoad.TestExplorerVisibleOnStartup;
            TodoExplorerVisibleOnStartup = toLoad.TodoExplorerVisibleOnStartup;
        }
    }
}
