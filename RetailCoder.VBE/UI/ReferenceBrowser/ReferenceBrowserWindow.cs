using System.Windows.Forms;

namespace Rubberduck.UI.ReferenceBrowser
{
    public partial class ReferenceBrowserWindow : Form
    {
        public ReferenceBrowserWindow(ReferenceBrowserViewModel viewModel)
        {
            InitializeComponent();

            referenceBrowser.DataContext = viewModel;
            viewModel.CloseWindow += ViewModel_CloseWindow;
        }

        private void ViewModel_CloseWindow(object sender, System.EventArgs e)
        {
            Close();
        }
    }
}
