using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveToFolder;

namespace Rubberduck.UI.Refactorings.MoveToFolder
{
    public partial class MoveMultipleToFolderView : IRefactoringView<MoveMultipleToFolderModel>
    {
        public MoveMultipleToFolderView()
        {
            InitializeComponent();

            Loaded += (o, e) =>
                {
                    MoveToFolderTextBox.Focus();
                    MoveToFolderTextBox.SelectAll();
                };
        }
    }
}
