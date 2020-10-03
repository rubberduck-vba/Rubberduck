using Rubberduck.Refactorings.MoveToFolder;
using Rubberduck.Resources;

namespace Rubberduck.UI.Refactorings.MoveToFolder
{
    internal class MoveMultipleToFolderPresenter : RefactoringPresenterBase<MoveMultipleToFolderModel>, IMoveMultipleToFolderPresenter
    {
        private static readonly DialogData DialogData = DialogData.Create(RubberduckUI.MoveToFolderDialog_Caption, 164, 684);

        public MoveMultipleToFolderPresenter(MoveMultipleToFolderModel model, IRefactoringDialogFactory dialogFactory) :
            base(DialogData, model, dialogFactory)
        {}
    }
}

