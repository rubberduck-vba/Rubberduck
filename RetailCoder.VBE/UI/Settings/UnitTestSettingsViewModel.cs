using NLog;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Settings
{
    public class UnitTestSettingsViewModel : SettingsViewModelBase, ISettingsViewModel
    {
        public UnitTestSettingsViewModel(Configuration config)
        {
            BindingMode = config.UserSettings.UnitTestSettings.BindingMode;
            AssertMode = config.UserSettings.UnitTestSettings.AssertMode;
            ModuleInit = config.UserSettings.UnitTestSettings.ModuleInit;
            MethodInit = config.UserSettings.UnitTestSettings.MethodInit;
            DefaultTestStubInNewModule = config.UserSettings.UnitTestSettings.DefaultTestStubInNewModule;
            ExportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ExportSettings());
            ImportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ImportSettings());
        }

        #region Properties

        private BindingMode _bindingMode;
        public BindingMode BindingMode
        {
            get { return _bindingMode; }
            set
            {
                if (_bindingMode != value)
                {
                    _bindingMode = value;
                    OnPropertyChanged();
                }
            }
        }

        private AssertMode _assertMode;
        public AssertMode AssertMode
        {
            get { return _assertMode; }
            set
            {
                if (_assertMode != value)
                {
                    _assertMode = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _moduleInit;
        public bool ModuleInit
        {
            get { return _moduleInit; }
            set
            {
                if (_moduleInit != value)
                {
                    _moduleInit = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _methodInit;
        public bool MethodInit
        {
            get { return _methodInit; }
            set
            {
                if (_methodInit != value)
                {
                    _methodInit = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _defaultTestStubInNewModule;
        public bool DefaultTestStubInNewModule
        {
            get { return _defaultTestStubInNewModule; }
            set
            {
                if (_defaultTestStubInNewModule != value)
                {
                    _defaultTestStubInNewModule = value;
                    OnPropertyChanged();
                }
            }
        }

        #endregion

        public void UpdateConfig(Configuration config)
        {
            config.UserSettings.UnitTestSettings.BindingMode = BindingMode;
            config.UserSettings.UnitTestSettings.AssertMode = AssertMode;
            config.UserSettings.UnitTestSettings.ModuleInit = ModuleInit;
            config.UserSettings.UnitTestSettings.MethodInit = MethodInit;
            config.UserSettings.UnitTestSettings.DefaultTestStubInNewModule = DefaultTestStubInNewModule;
        }

        public void SetToDefaults(Configuration config)
        {
            TransferSettingsToView(config.UserSettings.UnitTestSettings);
        }

        private void TransferSettingsToView(Rubberduck.Settings.UnitTestSettings toLoad)
        {
            BindingMode = toLoad.BindingMode;
            AssertMode = toLoad.AssertMode;
            ModuleInit = toLoad.ModuleInit;
            MethodInit = toLoad.MethodInit;
            DefaultTestStubInNewModule = toLoad.DefaultTestStubInNewModule;
        }

        private void ImportSettings()
        {
            using (var dialog = new OpenFileDialog
            {
                Filter = RubberduckUI.DialogMask_XmlFilesOnly,
                Title = RubberduckUI.DialogCaption_LoadUnitTestSettings
            })
            {
                dialog.ShowDialog();
                if (string.IsNullOrEmpty(dialog.FileName)) return;
                var service = new XmlPersistanceService<Rubberduck.Settings.UnitTestSettings> { FilePath = dialog.FileName };
                var loaded = service.Load(new Rubberduck.Settings.UnitTestSettings());
                TransferSettingsToView(loaded);
            }
        }

        private void ExportSettings()
        {
            using (var dialog = new SaveFileDialog
            {
                Filter = RubberduckUI.DialogMask_XmlFilesOnly,
                Title = RubberduckUI.DialogCaption_SaveUnitTestSettings
            })
            {
                dialog.ShowDialog();
                if (string.IsNullOrEmpty(dialog.FileName)) return;
                var service = new XmlPersistanceService<Rubberduck.Settings.UnitTestSettings> { FilePath = dialog.FileName };
                service.Save(new Rubberduck.Settings.UnitTestSettings
                {
                    BindingMode = BindingMode,
                    AssertMode = AssertMode,
                    ModuleInit = ModuleInit,
                    MethodInit = MethodInit,
                    DefaultTestStubInNewModule = DefaultTestStubInNewModule
                });
            }
        }
    }
}
