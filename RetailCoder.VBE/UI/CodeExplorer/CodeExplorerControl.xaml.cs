using System.Windows;
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
            if (ViewModel != null && ViewModel.OpenCommand.CanExecute(ViewModel.SelectedItem))
            {
                ViewModel.OpenCommand.Execute(ViewModel.SelectedItem);
            }
            e.Handled = true;
        }

        private void TreeView_OnMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            ((TreeViewItem)sender).IsSelected = true;
            e.Handled = true;
        }

        private void SearchBox_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            ViewModel.Search(ViewModel.Projects, ((TextBox)sender).Text);
        }

        private void SearchIcon_OnMouseDown(object sender, MouseButtonEventArgs e)
        {
            SearchBox.Focus();
        }

        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            SearchBox.Text = string.Empty;
            SearchBox.Focus();
        }
    }
}
