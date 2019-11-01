namespace Rubberduck.UI.Inspections
{
    /// <summary>
    /// Interaction logic for InspectionResultsControl.xaml
    /// </summary>
    public partial class InspectionResultsControl
    {
        private InspectionResultsViewModel ViewModel => DataContext as InspectionResultsViewModel;

        public InspectionResultsControl()
        {
            InitializeComponent();
        }

        private void InspectionResultsGrid_RequestBringIntoView(object sender, System.Windows.RequestBringIntoViewEventArgs e)
        {
            e.Handled = true;
        }
    }
}
