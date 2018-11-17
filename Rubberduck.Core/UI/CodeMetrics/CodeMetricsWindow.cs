using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;
using Rubberduck.Resources;
using Rubberduck.CodeAnalysis.CodeMetrics;

namespace Rubberduck.UI.CodeMetrics
{
    [ExcludeFromCodeCoverage]
    public sealed partial class CodeMetricsWindow : UserControl, IDockableUserControl
    {
        private const string ClassId = "C5318B5A-172F-417C-88E3-B377CDA2D809";
        string IDockableUserControl.ClassId => ClassId;
        string IDockableUserControl.Caption => RubberduckUI.CodeMetricsDockablePresenter_Caption;

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
