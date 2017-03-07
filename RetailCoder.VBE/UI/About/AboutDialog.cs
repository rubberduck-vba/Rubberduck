using System.Windows.Forms;

namespace Rubberduck.UI.About
{
    public partial class AboutDialog : Form
    {
        public AboutDialog()
        {
            InitializeComponent();

            // todo: inject these dependencies?
            ViewModel = new AboutControlViewModel(new VersionCheck.VersionCheck());
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
