using System.Windows;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveFolder;

namespace Rubberduck.UI.Refactorings.MoveFolder
{
    public partial class MoveMultipleFoldersView : IRefactoringView<MoveMultipleFoldersModel>
    {
        public MoveMultipleFoldersView()
        {
            InitializeComponent();

            Loaded += AfterLoadHandler;
        }

        private void AfterLoadHandler(object sender, RoutedEventArgs e)
        {
            MoveToFolderTextBox.Focus();
            MoveToFolderTextBox.SelectAll();
            Loaded -= AfterLoadHandler;
        }
    }
}
