using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;
using Rubberduck.CodeAnalysis;
using Rubberduck.CodeAnalysis.CodeMetrics;

namespace Rubberduck.UI.CodeMetrics
{
    [ExcludeFromCodeCoverage]
    public sealed partial class CodeMetricsWindow : UserControl, IDockableUserControl
    {
        private const string ClassId = "C5318B5A-172F-417C-88E3-B377CDA2D809";
        string IDockableUserControl.ClassId => ClassId;
        string IDockableUserControl.Caption => CodeAnalysisUI.CodeMetricsDockablePresenter_Caption;

        private CodeMetricsWindow()
        {
            InitializeComponent();
        }

        public CodeMetricsWindow(CodeMetricsViewModel viewModel) : this()
        {
            ViewModel = viewModel;
            codeMetricsControl1.DataContext = ViewModel;
        }

        public CodeMetricsViewModel ViewModel { get; }
    }
}
