namespace Rubberduck.UI.UnitTesting
{
    /// <summary>
    /// Interaction logic for TestExplorerControl.xaml
    /// </summary>
    public partial class TestExplorerControl
    {
        public TestExplorerControl()
        {
            InitializeComponent();
        }

        private void TestGrid_RequestBringIntoView(object sender, System.Windows.RequestBringIntoViewEventArgs e)
        {
            e.Handled = true;
        }
    }
}
