
namespace Rubberduck.UI.ToDoItems
{
    /// <summary>
    /// Interaction logic for ToDoExplorerControl.xaml
    /// </summary>
    public partial class ToDoExplorerControl
    {
        private ToDoExplorerViewModel ViewModel { get { return DataContext as ToDoExplorerViewModel; } }

        public ToDoExplorerControl()
        {
            InitializeComponent();
            Loaded += ToDoExplorerControl_Loaded;
        }

        void ToDoExplorerControl_Loaded(object sender, System.Windows.RoutedEventArgs e)
        {
            if (ViewModel != null && ViewModel.RefreshCommand.CanExecute(null))
            {
                ViewModel.RefreshCommand.Execute(null);
            }
        }
    }
}
