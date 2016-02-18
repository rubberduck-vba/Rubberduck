using System.Windows;
using System.Windows.Controls;

namespace Rubberduck.UI.CodeInspections
{
    /// <summary>
    /// Interaction logic for InspectionResultsControl.xaml
    /// </summary>
    public partial class InspectionResultsControl : UserControl
    {
        private InspectionResultsViewModel ViewModel { get { return DataContext as InspectionResultsViewModel; } }

        public InspectionResultsControl()
        {
            InitializeComponent();
            Loaded += InspectionResultsControl_Loaded;
        }

        private void InspectionResultsControl_Loaded(object sender, RoutedEventArgs e)
        {
            if (ViewModel != null && ViewModel.CanRefresh)
            {
                ViewModel.RefreshCommand.Execute(null);
            }
        }
    }
}
