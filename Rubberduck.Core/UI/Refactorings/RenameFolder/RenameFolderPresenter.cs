using Rubberduck.Refactorings.RenameFolder;
using Rubberduck.Refactorings;

namespace Rubberduck.UI.Refactorings.RenameFolder
{
    public class RenameFolderPresenter : RefactoringPresenterBase<RenameFolderModel>, IRenameFolderPresenter
    {
        private static readonly DialogData DialogData = DialogData.Create(RefactoringsUI.RenameDialog_Caption, 164, 684);

        public RenameFolderPresenter(RenameFolderModel model, IRefactoringDialogFactory dialogFactory) :
            base(DialogData, model, dialogFactory)
        { }
    }
}