using System.Windows.Forms;
using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    public partial class SettingsForm : Form
    {
        private readonly IGeneralConfigService _configService;

        public SettingsForm()
        {
            InitializeComponent();
        }

        public SettingsForm(IGeneralConfigService configService) : this()
        {
            _configService = configService;

            ViewModel = new SettingsControlViewModel(_configService);

            ViewModel.OnOKButtonClicked += ViewModel_OnOKButtonClicked;
            ViewModel.OnCancelButtonClicked += ViewModel_OnCancelButtonClicked;
        }

        void ViewModel_OnOKButtonClicked(object sender, System.EventArgs e)
        {
            Close();
        }

        void ViewModel_OnCancelButtonClicked(object sender, System.EventArgs e)
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
