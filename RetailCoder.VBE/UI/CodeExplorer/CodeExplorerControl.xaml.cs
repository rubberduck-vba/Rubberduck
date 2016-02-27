using System.Windows.Controls;
using System.Windows.Input;
using Rubberduck.Navigation.CodeExplorer;

namespace Rubberduck.UI.CodeExplorer
{
    /// <summary>
    /// Interaction logic for CodeExplorerControl.xaml
    /// </summary>
    public partial class CodeExplorerControl : UserControl
    {
        public CodeExplorerControl()
        {
            InitializeComponent();
        }

        private CodeExplorerViewModel ViewModel { get { return DataContext as CodeExplorerViewModel; } }

        private void TreeView_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ViewModel == null || ViewModel.SelectedItem == null)
            {
                return;
            }

            var selectedResult = ViewModel.SelectedItem as CodeExplorerItemViewModel;
            if (selectedResult == null || !selectedResult.QualifiedSelection.HasValue)
            {
                return;
            }

            var arg = selectedResult.QualifiedSelection.Value.GetNavitationArgs();
            ViewModel.NavigateCommand.Execute(arg);
        }
    }
}
