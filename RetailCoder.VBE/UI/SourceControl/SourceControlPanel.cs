using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;

namespace Rubberduck.UI.SourceControl
{
    [ExcludeFromCodeCoverage]
    public partial class SourceControlPanel : UserControl, IDockableUserControl
    {
        private SourceControlPanel()
        {
            InitializeComponent();
        }

        public SourceControlPanel(SourceControlViewViewModel viewModel) : this()
        {
            _viewModel = viewModel;
            SourceControlPanelControl.DataContext = viewModel;
        }

        public string ClassId
        {
            get { return "19A32FC9-4902-4385-9FE7-829D4F9C441D"; }
        }

        public string Caption
        {
            get { return RubberduckUI.SourceControlPanel_Caption; }
        }

        private readonly SourceControlViewViewModel _viewModel;
        public SourceControlViewViewModel ViewModel
        {
            get { return _viewModel; }
        }
    }
}
