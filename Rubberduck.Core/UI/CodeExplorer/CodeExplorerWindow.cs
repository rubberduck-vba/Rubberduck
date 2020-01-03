using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Resources.CodeExplorer;

namespace Rubberduck.UI.CodeExplorer
{
    [ExcludeFromCodeCoverage]
    public sealed partial class CodeExplorerWindow : UserControl, IDockableUserControl
    {
        private const string ClassId = "C5318B59-172F-417C-88E3-B377CDA2D809";
        string IDockableUserControl.ClassId => ClassId;
        string IDockableUserControl.Caption => CodeExplorerUI.CodeExplorerDockablePresenter_Caption;

        private CodeExplorerWindow()
        {
            InitializeComponent();
        }

        public CodeExplorerWindow(CodeExplorerViewModel viewModel) : this()
        {
            ViewModel = viewModel;
            codeExplorerControl1.DataContext = ViewModel;
        }
        public CodeExplorerViewModel ViewModel { get; }
    }
}
