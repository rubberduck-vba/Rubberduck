using System.Windows.Forms;
using Rubberduck.Resources.UnitTesting;

namespace Rubberduck.UI.UnitTesting
{
    public partial class TestExplorerWindow : UserControl, IDockableUserControl
    {
        private TestExplorerWindow()
        {
            InitializeComponent();
        }

        public TestExplorerWindow(TestExplorerViewModel viewModel) : this()
        {
            _viewModel = viewModel;
            wpfTestExplorerControl.DataContext = _viewModel;
        }

        private readonly TestExplorerViewModel _viewModel;
        public TestExplorerViewModel ViewModel
        {
            get { return _viewModel; }
        }

        private readonly string RandomGuid = Guid.NewGuid().ToString();
        string IDockableUserControl.GuidIdentifier => RandomGuid;

        public string Caption => TestExplorer.TestExplorerWindow_Caption;
    }
}
