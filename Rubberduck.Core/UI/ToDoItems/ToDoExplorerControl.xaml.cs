using System;

namespace Rubberduck.UI.ToDoItems
{
    /// <summary>
    /// Interaction logic for ToDoExplorerControl.xaml
    /// </summary>
    public partial class ToDoExplorerControl
    {
        private ToDoExplorerViewModel ViewModel => DataContext as ToDoExplorerViewModel;

        public ToDoExplorerControl()
        {
            InitializeComponent();
            Loaded += ToDoExplorerControl_Loaded;
        }

        private void ToDoExplorerControl_Loaded(object sender, System.Windows.RoutedEventArgs e)
        {
            if (ViewModel != null && ViewModel.RefreshCommand.CanExecute(null))
            {
                ViewModel.RefreshCommand.Execute(null);
            }
        }

        private void GroupingGrid_ColumnReordered(object sender, System.Windows.Controls.DataGridColumnEventArgs e)
        {
            ViewModel.UpdateColumnHeaderInformation(MainGrid.Columns);
        }
    }
}
