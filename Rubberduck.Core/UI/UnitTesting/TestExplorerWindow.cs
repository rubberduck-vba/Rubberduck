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
        public TestExplorerViewModel ViewModel => _viewModel;

        public string ClassId => "9CF1392A-2DC9-48A6-AC0B-E601A9802608";

        public string Caption => TestExplorer.TestExplorerWindow_Caption;
    }
}
