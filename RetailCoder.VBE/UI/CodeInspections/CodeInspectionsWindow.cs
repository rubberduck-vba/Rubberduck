using System.Windows.Forms;

namespace Rubberduck.UI.CodeInspections
{
    public partial class CodeInspectionsWindow : UserControl, IDockableUserControl
    {
        private const string ClassId = "D3B2A683-9856-4246-BDC8-6B0795DC875B";
        string IDockableUserControl.ClassId { get { return ClassId; } }
        string IDockableUserControl.Caption { get { return RubberduckUI.CodeInspections; } }
        
        public CodeInspectionsWindow()
        {
            InitializeComponent();
        }

        private InspectionResultsViewModel _viewModel;
        public InspectionResultsViewModel ViewModel
        {
            get { return _viewModel; }
            set
            {
                _viewModel = value;
                wpfInspectionResultsControl.DataContext = _viewModel;
            }
        }
    }
}
