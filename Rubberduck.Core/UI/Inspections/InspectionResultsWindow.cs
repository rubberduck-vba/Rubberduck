using System;
using System.Windows.Forms;
using Rubberduck.Resources;

namespace Rubberduck.UI.Inspections
{
    public partial class InspectionResultsWindow : UserControl, IDockableUserControl
    {
        private readonly string RandomGuid = Guid.NewGuid().ToString();
        string IDockableUserControl.GuidIdentifier => RandomGuid;
        string IDockableUserControl.Caption { get { return RubberduckUI.CodeInspections; } }
        
        private InspectionResultsWindow()
        {
            InitializeComponent();
        }

        public InspectionResultsWindow(InspectionResultsViewModel viewModel) : this()
        {
            _viewModel = viewModel;
            wpfInspectionResultsControl.DataContext = _viewModel;
        }

        private readonly InspectionResultsViewModel _viewModel;
        public InspectionResultsViewModel ViewModel
        {
            get { return _viewModel; }
        }
    }
}
