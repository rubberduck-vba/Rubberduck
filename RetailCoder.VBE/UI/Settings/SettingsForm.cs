using System.Windows.Forms;

namespace Rubberduck.UI.Settings
{
    public partial class SettingsForm : Form
    {
        public SettingsForm()
        {
            InitializeComponent();

            ViewModel = new SettingsControlViewModel();

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
