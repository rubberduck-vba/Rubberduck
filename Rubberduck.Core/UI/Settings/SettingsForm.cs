using System.Windows.Forms;
using Rubberduck.Settings;
using Rubberduck.Interaction;
using Rubberduck.SettingsProvider;
using Rubberduck.CodeAnalysis.Settings;

namespace Rubberduck.UI.Settings
{
    public interface ISettingsFormFactory
    {
        SettingsForm Create();
        SettingsForm Create(SettingsViews activeView);
        void Release(SettingsForm form);
    }

    public interface ISettingsViewModelFactory
    {
        ISettingsViewModel<TSettings> Create<TSettings>(Configuration config) where TSettings : class, new();
        ISettingsViewModel<TSettings> Create<TSettings>() where TSettings : class, new();
        void Release<TSettings>(ISettingsViewModel<TSettings> viewModel) where TSettings : class, new();
    }

    public partial class SettingsForm : Form
    {
        public SettingsForm()
        {
            InitializeComponent();
        }

        public SettingsForm(IConfigurationService<Configuration> configService, 
            IMessageBox messageBox, 
            ISettingsViewModelFactory viewModelFactory,
            SettingsViews activeView = SettingsViews.GeneralSettings) 
            : this()
        {
            var config = configService.Read();

            ViewModel = new SettingsControlViewModel(messageBox, configService,
                config,
                new GeneralSettingsView
                {
                    Control = new GeneralSettings(viewModelFactory.Create<Rubberduck.Settings.GeneralSettings>(config)),
                    View = SettingsViews.GeneralSettings
                },
                new SettingsView
                {
                    Control = new TodoSettings(viewModelFactory.Create<ToDoListSettings>(config)),
                    View = SettingsViews.TodoSettings
                },
                new SettingsView
                {
                    Control = new InspectionSettings(viewModelFactory.Create<CodeInspectionSettings>(config)),
                    View = SettingsViews.InspectionSettings
                },
                new SettingsView
                {
                    Control = new UnitTestSettings(viewModelFactory.Create<Rubberduck.UnitTesting.Settings.UnitTestSettings>(config)),
                    View = SettingsViews.UnitTestSettings
                },
                new SettingsView
                {
                    Control = new IndenterSettings(viewModelFactory.Create<SmartIndenter.IndenterSettings>(config)),
                    View = SettingsViews.IndenterSettings
                },
                new SettingsView
                {
                    Control = new AutoCompleteSettings(viewModelFactory.Create<Rubberduck.Settings.AutoCompleteSettings>(config)),
                    View = SettingsViews.AutoCompleteSettings
                },
                new SettingsView
                {
                    Control = new WindowSettings(viewModelFactory.Create<Rubberduck.Settings.WindowSettings>(config)),
                    View = SettingsViews.WindowSettings
                },
                new SettingsView
                {
                    Control = new AddRemoveReferencesUserSettings(viewModelFactory.Create<ReferenceSettings>()),
                    View = SettingsViews.ReferenceSettings
                },
                activeView);

            ViewModel.OnWindowClosed += ViewModel_OnWindowClosed;
        }

        void ViewModel_OnWindowClosed(object sender, System.EventArgs e)
        {
            Close();
        }

        private SettingsControlViewModel _viewModel;
        private SettingsControlViewModel ViewModel
        {
            get { return _viewModel; }
            set
            {
                _viewModel = value;
                SettingsControl.DataContext = _viewModel;
            }
        }
    }
}
