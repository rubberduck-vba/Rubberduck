using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Rename;

namespace Rubberduck.UI.Refactorings.Rename
{
    internal class RenamePresenter : RefactoringPresenterBase<RenameModel, RenameDialog, RenameView, RenameViewModel>, IRenamePresenter
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

