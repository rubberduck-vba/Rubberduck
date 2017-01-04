using NLog;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Settings
{
    public class WindowSettingsViewModel : SettingsViewModelBase, ISettingsViewModel
    {
        public WindowSettingsViewModel(Configuration config)
        {
            CodeExplorerVisibleOnStartup = config.UserSettings.WindowSettings.CodeExplorerVisibleOnStartup;
            CodeInspectionsVisibleOnStartup = config.UserSettings.WindowSettings.CodeInspectionsVisibleOnStartup;
            SourceControlVisibleOnStartup = config.UserSettings.WindowSettings.SourceControlVisibleOnStartup;
            TestExplorerVisibleOnStartup = config.UserSettings.WindowSettings.TestExplorerVisibleOnStartup;
            TodoExplorerVisibleOnStartup = config.UserSettings.WindowSettings.TodoExplorerVisibleOnStartup;

            ExportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ExportSettings());
            ImportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ImportSettings());
        }

        #region Properties

        private bool _codeExplorerVisibleOnStartup;
        public bool CodeExplorerVisibleOnStartup
        {
            get { return _codeExplorerVisibleOnStartup; }
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
            get { return _codeInspectionsVisibleOnStartup; }
            set
            {
                if (_codeInspectionsVisibleOnStartup != value)
                {
                    _codeInspectionsVisibleOnStartup = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _sourceControlVisibleOnStartup;
        public bool SourceControlVisibleOnStartup
        {
            get { return _sourceControlVisibleOnStartup; }
            set
            {
                if (_sourceControlVisibleOnStartup != value)
                {
                    _sourceControlVisibleOnStartup = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _testExplorerVisibleOnStartup;
        public bool TestExplorerVisibleOnStartup
        {
            get { return _testExplorerVisibleOnStartup; }
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
            get { return _todoExplorerVisibleOnStartup; }
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
            config.UserSettings.WindowSettings.SourceControlVisibleOnStartup = SourceControlVisibleOnStartup;
            config.UserSettings.WindowSettings.TestExplorerVisibleOnStartup = TestExplorerVisibleOnStartup;
            config.UserSettings.WindowSettings.TodoExplorerVisibleOnStartup = TodoExplorerVisibleOnStartup;
        }

        public void SetToDefaults(Configuration config)
        {
            TransferSettingsToView(config.UserSettings.WindowSettings);
        }

        private void TransferSettingsToView(Rubberduck.Settings.WindowSettings toLoad)
        {
            CodeExplorerVisibleOnStartup = toLoad.CodeExplorerVisibleOnStartup;
            CodeInspectionsVisibleOnStartup = toLoad.CodeInspectionsVisibleOnStartup;
            SourceControlVisibleOnStartup = toLoad.SourceControlVisibleOnStartup;
            TestExplorerVisibleOnStartup = toLoad.TestExplorerVisibleOnStartup;
            TodoExplorerVisibleOnStartup = toLoad.TodoExplorerVisibleOnStartup;
        }

        private void ImportSettings()
        {
            using (var dialog = new OpenFileDialog
            {
                Filter = RubberduckUI.DialogMask_XmlFilesOnly,
                Title = RubberduckUI.DialogCaption_LoadWindowSettings
            })
            {
                dialog.ShowDialog();
                if (string.IsNullOrEmpty(dialog.FileName)) return;
                var service = new XmlPersistanceService<Rubberduck.Settings.WindowSettings> { FilePath = dialog.FileName };
                var loaded = service.Load(new Rubberduck.Settings.WindowSettings());
                TransferSettingsToView(loaded);
            }
        }

        private void ExportSettings()
        {
            using (var dialog = new SaveFileDialog
            {
                Filter = RubberduckUI.DialogMask_XmlFilesOnly,
                Title = RubberduckUI.DialogCaption_SaveWindowSettings
            })
            {
                dialog.ShowDialog();
                if (string.IsNullOrEmpty(dialog.FileName)) return;
                var service = new XmlPersistanceService<Rubberduck.Settings.WindowSettings> { FilePath = dialog.FileName };
                service.Save(new Rubberduck.Settings.WindowSettings()
                {
                    CodeExplorerVisibleOnStartup = CodeExplorerVisibleOnStartup,
                    CodeInspectionsVisibleOnStartup = CodeInspectionsVisibleOnStartup,
                    SourceControlVisibleOnStartup = SourceControlVisibleOnStartup,
                    TestExplorerVisibleOnStartup = TestExplorerVisibleOnStartup,
                    TodoExplorerVisibleOnStartup = TodoExplorerVisibleOnStartup
                });
            }
        }
    }
}
