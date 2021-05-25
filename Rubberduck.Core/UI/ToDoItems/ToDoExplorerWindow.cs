using System;
using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;
using Rubberduck.Resources.ToDoExplorer;

namespace Rubberduck.UI.ToDoItems
{
    [ExcludeFromCodeCoverage]
    public sealed partial class ToDoExplorerWindow : UserControl, IDockableUserControl
    {
        private const string ClassId = "8B071EDA-2C9C-4009-9A22-A1958BF98B28"; // todo get from Resources.Registration?
        string IDockableUserControl.ClassId => ClassId;
        string IDockableUserControl.Caption => Rubberduck.Resources.ToDoExplorer.ToDoExplorerUI.TodoExplorer_Caption;

        private ToDoExplorerWindow()
        {
            InitializeComponent();
        }

        public ToDoExplorerWindow(ToDoExplorerViewModel viewModel) : this()
        {
            ViewModel = viewModel;
            TodoExplorerControl.DataContext = ViewModel;
            viewModel.UpdateColumnHeaderInformationToMatchCached(TodoExplorerControl.MainGrid.Columns);
        }
        public ToDoExplorerViewModel ViewModel { get; }
    }
}
