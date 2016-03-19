using System.Windows.Forms;

namespace Rubberduck.UI.About
{
    public partial class AboutDialog : Form
    {
        public AboutDialog()
        {
            InitializeComponent();

            ViewModel = new AboutControlViewModel();
        }

        private AboutControlViewModel _viewModel;
        private AboutControlViewModel ViewModel
        {
            get { return _viewModel; }
            set
            {
                _viewModel = value;
                AboutControl.DataContext = _viewModel;
            }
        }
    }
}
