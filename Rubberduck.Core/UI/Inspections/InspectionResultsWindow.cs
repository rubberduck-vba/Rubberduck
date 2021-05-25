using System.Windows.Forms;
using Rubberduck.CodeAnalysis;

namespace Rubberduck.UI.Inspections
{
    public sealed partial class InspectionResultsWindow : UserControl, IDockableUserControl
    {
        private const string ClassId = "D3B2A683-9856-4246-BDC8-6B0795DC875B";
        string IDockableUserControl.ClassId => ClassId;
        string IDockableUserControl.Caption => CodeAnalysisUI.CodeInspections;

        private InspectionResultsWindow()
        {
            InitializeComponent();
        }

        public InspectionResultsWindow(InspectionResultsViewModel viewModel) : this()
        {
            ViewModel = viewModel;
            wpfInspectionResultsControl.DataContext = ViewModel;
        }

        public InspectionResultsViewModel ViewModel { get; }
    }
}
