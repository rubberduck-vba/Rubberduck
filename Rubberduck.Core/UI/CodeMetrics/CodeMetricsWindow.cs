using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;
using Rubberduck.Resources;
using Rubberduck.CodeAnalysis.CodeMetrics;
using System;

namespace Rubberduck.UI.CodeMetrics
{
    [ExcludeFromCodeCoverage]
    public partial class CodeMetricsWindow : UserControl, IDockableUserControl
    {
        private readonly string RandomGuid = Guid.NewGuid().ToString();
        string IDockableUserControl.GuidIdentifier => RandomGuid;
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
