using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Refactorings;
using Rubberduck.UI.Refactorings.Rename;

namespace Rubberduck.Refactorings.Rename
{
    public class RenamePresenter : RefactoringPresenterBase<RenameModel, RenameDialog, RenameView, RenameViewModel>
    {
        public RenamePresenter(RenameModel model,
            IRefactoringDialogFactory<RenameModel, RenameView, RenameViewModel, RenameDialog> dialogFactory) : base(
            model, dialogFactory)
        { }
        
        public override RenameModel Show()
        {
            return Model.Target == null ? null : Show(Model.Target);
        }

        public RenameModel Show(Declaration target)
        {
            if (null == target)
            {
                return null;
            }

            Model.Target = target;
            ViewModel.Target = target;

            Show();

            if (DialogResult != RefactoringDialogResult.Execute)
            {
                return null;
            }

            Model.NewName = ViewModel.NewName;
            return Model;
        }
    }
}

