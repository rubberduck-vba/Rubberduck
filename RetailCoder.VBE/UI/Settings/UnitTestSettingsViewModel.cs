using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    public class UnitTestSettingsViewModel : ViewModelBase, ISettingsViewModel
    {
        public UnitTestSettingsViewModel(Configuration config)
        {
            BindingMode = config.UserSettings.UnitTestSettings.BindingMode;
            AssertMode = config.UserSettings.UnitTestSettings.AssertMode;
            ModuleInit = config.UserSettings.UnitTestSettings.ModuleInit;
            MethodInit = config.UserSettings.UnitTestSettings.MethodInit;
            DefaultTestStubInNewModule = config.UserSettings.UnitTestSettings.DefaultTestStubInNewModule;
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
            BindingMode = config.UserSettings.UnitTestSettings.BindingMode;
            AssertMode = config.UserSettings.UnitTestSettings.AssertMode;
            ModuleInit = config.UserSettings.UnitTestSettings.ModuleInit;
            MethodInit = config.UserSettings.UnitTestSettings.MethodInit;
            DefaultTestStubInNewModule = config.UserSettings.UnitTestSettings.DefaultTestStubInNewModule;
        }
    }
}
