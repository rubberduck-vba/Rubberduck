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

        private void AboutDialog_Load(object sender, System.EventArgs e)
        {

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
