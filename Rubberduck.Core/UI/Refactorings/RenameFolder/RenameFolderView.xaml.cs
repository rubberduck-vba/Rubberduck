using System.Windows;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.RenameFolder;

namespace Rubberduck.UI.Refactorings.RenameFolder
{
    public partial class RenameFolderView : IRefactoringView<RenameFolderModel>
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
