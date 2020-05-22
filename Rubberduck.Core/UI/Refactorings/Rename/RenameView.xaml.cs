using System.Windows;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Rename;

namespace Rubberduck.UI.Refactorings.Rename
{
    public partial class RenameView : IRefactoringView<RenameModel>
    {
        public RenameView()
        {
            InitializeComponent();

            Loaded += AfterLoadHandler;
        }

        private void AfterLoadHandler(object sender, RoutedEventArgs e)
        {
            RenameTextBox.Focus();
            RenameTextBox.SelectAll();
            Loaded -= AfterLoadHandler;
        }
    }
}
