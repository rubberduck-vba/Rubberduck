using System.Windows;
using Rubberduck.Refactorings;

namespace Rubberduck.UI.Refactorings.RenameFolder
{
    public partial class RenameFolderView : IRefactoringView<RenameFolderView>
    {
        public RenameFolderView()
        {
            InitializeComponent();

            Loaded += AfterLoadHandler;
        }

        private void AfterLoadHandler(object sender, RoutedEventArgs e)
        {
            RenameFolderTextBox.Focus();
            RenameFolderTextBox.SelectAll();
            Loaded -= AfterLoadHandler;
        }
    }
}
