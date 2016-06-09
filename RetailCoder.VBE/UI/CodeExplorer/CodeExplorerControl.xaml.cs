using System.Windows.Controls;
using System.Windows.Input;
using Rubberduck.Navigation.CodeExplorer;

namespace Rubberduck.UI.CodeExplorer
{
    /// <summary>
    /// Interaction logic for CodeExplorerControl.xaml
    /// </summary>
    public partial class CodeExplorerControl
    {
        public CodeExplorerControl()
        {
            InitializeComponent();
        }

        private CodeExplorerViewModel ViewModel { get { return DataContext as CodeExplorerViewModel; } }

        private void TreeView_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ViewModel != null && ViewModel.NavigateCommand.CanExecute(ViewModel.SelectedItem))
            {
                ViewModel.NavigateCommand.Execute(ViewModel.SelectedItem);
            }
        }

        private void OnPreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            ((TreeViewItem)sender).IsSelected = true;
        }
    }
}
