using System.Windows.Forms;
using Rubberduck.Settings;
using Rubberduck.Common;

namespace Rubberduck.UI.Settings
{
    public partial class SettingsForm : Form
    {
        public SettingsForm()
        {
            InitializeComponent();
        }

        public SettingsForm(IGeneralConfigService configService, IOperatingSystem operatingSystem, SettingsViews activeView = SettingsViews.GeneralSettings) : this()
        {
            var config = configService.LoadConfiguration();

            ViewModel = new SettingsControlViewModel(configService,
                config,
                new SettingsView
                {
                    Control = new GeneralSettings(new GeneralSettingsViewModel(config, operatingSystem)),
                    View = SettingsViews.GeneralSettings
                },
                new SettingsView
                {
                    Control = new TodoSettings(new TodoSettingsViewModel(config)),
                    View = SettingsViews.TodoSettings
                },
                new SettingsView
                {
                    Control = new InspectionSettings(new InspectionSettingsViewModel(config)),
                    View = SettingsViews.InspectionSettings
                },
                new SettingsView
                {
                    Control = new UnitTestSettings(new UnitTestSettingsViewModel(config)),
                    View = SettingsViews.UnitTestSettings
                },
                new SettingsView
                {
                    Control = new IndenterSettings(new IndenterSettingsViewModel(config)),
                    View = SettingsViews.IndenterSettings
                },
                new SettingsView
                {
                    Control = new WindowSettings(new WindowSettingsViewModel(config)),
                    View = SettingsViews.WindowSettings
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
