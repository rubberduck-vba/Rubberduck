using System.Windows.Input;

namespace Rubberduck.UI.ToDoItems
{
    /// <summary>
    /// Interaction logic for ToDoExplorerControl.xaml
    /// </summary>
    public partial class ToDoExplorerControl
    {
        public ToDoExplorerControl()
        {
            InitializeComponent();
        }

        private void GroupingGridItem_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var viewModel = DataContext as ToDoExplorerViewModel;

            // this seems idiotic, but if you hold CTRL while you double-click an item
            // it both unselected the item and triggers the double-click, resulting in an NRE here
            if (viewModel != null && viewModel.SelectedToDo != null)
            {
                viewModel.NavigateToToDo.Execute(new NavigateCodeEventArgs(viewModel.SelectedToDo.GetSelection()));
            }
        }
    }
}
