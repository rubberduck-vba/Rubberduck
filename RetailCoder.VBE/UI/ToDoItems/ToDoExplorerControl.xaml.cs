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

            if (viewModel != null)
            {
                viewModel.NavigateToToDo.Execute(new NavigateCodeEventArgs(viewModel.SelectedToDo.GetSelection()));
            }
        }
    }
}
