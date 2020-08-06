using Rubberduck.VersionCheck;
using System.Windows.Forms;

namespace Rubberduck.UI.About
{
    public partial class AboutDialog : Form
    {
        public AboutDialog(IVersionCheck versionCheck, IWebNavigator web) : this()
        {
            ViewModel = new AboutControlViewModel(versionCheck, web);
        }

        public AboutDialog()
        {
            InitializeComponent();
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

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                this.Close();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
    }
}
