using Rubberduck.Refactorings.RenameFolder;
using Rubberduck.Resources;

namespace Rubberduck.UI.Refactorings.RenameFolder
{
    public class RenameFolderPresenter : RefactoringPresenterBase<RenameFolderModel>, IRenameFolderPresenter
    {
        private static readonly DialogData DialogData = DialogData.Create(RubberduckUI.RenameDialog_Caption, 164, 684);

        public RenameFolderPresenter(RenameFolderModel model, IRefactoringDialogFactory dialogFactory) :
            base(DialogData, model, dialogFactory)
        { }
    }
}