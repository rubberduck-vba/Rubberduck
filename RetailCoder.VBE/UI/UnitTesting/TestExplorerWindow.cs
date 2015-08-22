using System.Windows.Forms;

namespace Rubberduck.UI.UnitTesting
{
    public partial class TestExplorerWindow : UserControl, IDockableUserControl
    {
        public string ClassId
        {
            get { return "9CF1392A-2DC9-48A6-AC0B-E601A9802608"; }
        }

        public string Caption
        {
            get { return RubberduckUI.TestExplorerWindow_Caption; }
        }

        public TestExplorerWindow()
        {
            InitializeComponent();
        }

        public TestExplorerWindow(TestExplorerViewModel viewModel)
            : this()
        {
            ((TestExplorerControl)(wpfHost.Child)).DataContext = viewModel;
        }
    }
}
