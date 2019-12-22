using NLog;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.Command;
using Rubberduck.Resources.Settings;
using Rubberduck.UnitTesting.Settings;

namespace Rubberduck.UI.Settings
{
    public sealed class UnitTestSettingsViewModel : SettingsViewModelBase<Rubberduck.UnitTesting.Settings.UnitTestSettings>, ISettingsViewModel<Rubberduck.UnitTesting.Settings.UnitTestSettings>
    {
        public UnitTestSettingsViewModel(Configuration config, IConfigurationService<Rubberduck.UnitTesting.Settings.UnitTestSettings> service) 
            : base(service)
        {
            BindingMode = config.UserSettings.UnitTestSettings.BindingMode;
            AssertMode = config.UserSettings.UnitTestSettings.AssertMode;
            ModuleInit = config.UserSettings.UnitTestSettings.ModuleInit;
            MethodInit = config.UserSettings.UnitTestSettings.MethodInit;
            DefaultTestStubInNewModule = config.UserSettings.UnitTestSettings.DefaultTestStubInNewModule;
            ExportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(),
                _ => ExportSettings(new Rubberduck.UnitTesting.Settings.UnitTestSettings(BindingMode, AssertMode, ModuleInit,
                    MethodInit, DefaultTestStubInNewModule)));
            ImportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ImportSettings());
        }

        #region Properties

        private BindingMode _bindingMode;
        public BindingMode BindingMode
        {
            get => _bindingMode;
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
            get => _assertMode;
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
            get => _moduleInit;
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
            get => _methodInit;
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
            get => _defaultTestStubInNewModule;
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

        protected override string DialogLoadTitle => SettingsUI.DialogCaption_LoadUnitTestSettings;
        protected override string DialogSaveTitle => SettingsUI.DialogCaption_SaveUnitTestSettings;
        protected override void TransferSettingsToView(Rubberduck.UnitTesting.Settings.UnitTestSettings toLoad)
        {
            BindingMode = toLoad.BindingMode;
            AssertMode = toLoad.AssertMode;
            ModuleInit = toLoad.ModuleInit;
            MethodInit = toLoad.MethodInit;
            DefaultTestStubInNewModule = toLoad.DefaultTestStubInNewModule;
        }
    }
}
