using System.Windows.Forms;

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

        public string ClassId
        {
            get { return "9CF1392A-2DC9-48A6-AC0B-E601A9802608"; }
        }

        public string Caption
        {
            get { return RubberduckUI.TestExplorerWindow_Caption; }
        }
    }
}
