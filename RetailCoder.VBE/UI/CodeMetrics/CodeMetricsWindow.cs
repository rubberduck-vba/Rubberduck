using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;
using Rubberduck.Navigation.CodeMetrics;

namespace Rubberduck.UI.CodeMetrics
{
    [ExcludeFromCodeCoverage]
    public partial class CodeMetricsWindow : UserControl, IDockableUserControl
    {
        private const string ClassId = "C5318B59-172F-417C-88E3-B377CDA2D80A";
        string IDockableUserControl.ClassId { get { return ClassId; } }
        string IDockableUserControl.Caption { get { return RubberduckUI.CodeExplorerDockablePresenter_Caption; } }

        private CodeMetricsWindow()
        {
            InitializeComponent();
        }

        public CodeMetricsWindow(CodeMetricsViewModel viewModel) : this()
        {
            _viewModel = viewModel;
            codeMetricsControl1.DataContext = _viewModel;
        }

        private readonly CodeMetricsViewModel _viewModel;
        public CodeMetricsViewModel ViewModel
        {
            get { return _viewModel; }
        }
    }
}
