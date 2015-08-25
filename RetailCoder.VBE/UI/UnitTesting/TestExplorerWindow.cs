using System.Windows.Forms;

namespace Rubberduck.UI.UnitTesting
{
    public partial class TestExplorerWindow : UserControl, IDockableUserControl
    {
        public TestExplorerWindow()
        {
            InitializeComponent();
        }

        private TestExplorerViewModel _viewModel;
        public TestExplorerViewModel ViewModel
        {
            get { return _viewModel; }
            set
            {
                _viewModel = value;
                wpfTestExplorerControl.DataContext = _viewModel;
            }
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
