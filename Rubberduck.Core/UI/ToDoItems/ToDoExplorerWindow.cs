using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;
using Rubberduck.Resources;

namespace Rubberduck.UI.ToDoItems
{
    [ExcludeFromCodeCoverage]
    public partial class ToDoExplorerWindow : UserControl, IDockableUserControl
    {
        private readonly string RandomGuid = Guid.NewGuid().ToString();
        string IDockableUserControl.GuidIdentifier => RandomGuid;
        string IDockableUserControl.Caption { get { return RubberduckUI.TodoExplorer_Caption; } }

        private ToDoExplorerWindow()
        {
            InitializeComponent();
        }

        public ToDoExplorerWindow(ToDoExplorerViewModel viewModel) : this()
        {
            _viewModel = viewModel;
            TodoExplorerControl.DataContext = _viewModel;
        }

        private readonly ToDoExplorerViewModel _viewModel;
        public ToDoExplorerViewModel ViewModel
        {
            get { return _viewModel; }
        }
    }
}
