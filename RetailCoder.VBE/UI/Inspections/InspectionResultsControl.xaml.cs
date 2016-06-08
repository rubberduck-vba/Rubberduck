namespace Rubberduck.UI.Inspections
{
    /// <summary>
    /// Interaction logic for InspectionResultsControl.xaml
    /// </summary>
    public partial class InspectionResultsControl
    {
        private InspectionResultsViewModel ViewModel { get { return DataContext as InspectionResultsViewModel; } }

        public InspectionResultsControl()
        {
            InitializeComponent();
        }
    }
}
