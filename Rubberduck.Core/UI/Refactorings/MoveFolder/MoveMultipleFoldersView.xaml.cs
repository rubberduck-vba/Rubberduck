using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveFolder;

namespace Rubberduck.UI.Refactorings.MoveFolder
{
    public partial class MoveMultipleFoldersView : IRefactoringView<MoveMultipleFoldersModel>
    {
        public MoveMultipleFoldersView()
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
