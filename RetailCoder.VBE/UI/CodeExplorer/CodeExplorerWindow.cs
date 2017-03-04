using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;
using Rubberduck.Navigation.CodeExplorer;

namespace Rubberduck.UI.CodeExplorer
{
    [ExcludeFromCodeCoverage]
    public partial class CodeExplorerWindow : UserControl, IDockableUserControl
    {
        private const string ClassId = "C5318B59-172F-417C-88E3-B377CDA2D809";
        string IDockableUserControl.ClassId { get { return ClassId; } }
        string IDockableUserControl.Caption { get { return RubberduckUI.CodeExplorerDockablePresenter_Caption; } }

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
