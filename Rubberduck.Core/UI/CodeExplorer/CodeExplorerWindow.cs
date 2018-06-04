using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Resources.CodeExplorer;

namespace Rubberduck.UI.CodeExplorer
{
    [ExcludeFromCodeCoverage]
    public partial class CodeExplorerWindow : UserControl, IDockableUserControl
    {
        private readonly string RandomGuid = Guid.NewGuid().ToString();
        string IDockableUserControl.GuidIdentifier => RandomGuid;
        string IDockableUserControl.Caption { get { return CodeExplorerUI.CodeExplorerDockablePresenter_Caption; } }

        private CodeExplorerWindow()
        {
            InitializeComponent();
        }

        public CodeExplorerWindow(CodeExplorerViewModel viewModel) : this()
        {
            _viewModel = viewModel;
            codeExplorerControl1.DataContext = _viewModel;
        }

        private readonly CodeExplorerViewModel _viewModel;
        public CodeExplorerViewModel ViewModel
        {
            get { return _viewModel; }
        }
    }
}
