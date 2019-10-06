using System.Windows.Forms;
using Rubberduck.Resources.UnitTesting;

namespace Rubberduck.UI.UnitTesting
{
    internal partial class TestExplorerWindow : UserControl, IDockableUserControl
    {
        private TestExplorerWindow()
        {
            InitializeComponent();
        }

        public TestExplorerWindow(TestExplorerViewModel viewModel) : this()
        {
            ViewModel = viewModel;
            wpfTestExplorerControl.DataContext = ViewModel;
        }
        public TestExplorerViewModel ViewModel { get; }

        // FIXME bare ClassId... not good
        public string ClassId => "9CF1392A-2DC9-48A6-AC0B-E601A9802608";

        public string Caption => TestExplorer.TestExplorerWindow_Caption;
    }
}
