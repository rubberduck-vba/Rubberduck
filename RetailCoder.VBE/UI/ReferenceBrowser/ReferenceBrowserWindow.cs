using System.Windows.Forms;

namespace Rubberduck.UI.ReferenceBrowser
{
    public partial class ReferenceBrowserWindow : Form
    {
        public ReferenceBrowserWindow(ReferenceBrowserViewModel viewModel)
        {
            InitializeComponent();

            referenceBrowser.DataContext = viewModel;
        }
    }
}
