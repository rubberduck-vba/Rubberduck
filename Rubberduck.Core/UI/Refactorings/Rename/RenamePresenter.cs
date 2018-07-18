using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Rename;

namespace Rubberduck.UI.Refactorings.Rename
{
    internal class RenamePresenter : RefactoringPresenterBase<RenameModel, IRefactoringDialog<RenameModel, IRefactoringView<RenameModel>, RenameViewModel>, IRefactoringView<RenameModel>, RenameViewModel>, IRenamePresenter
    {
        public RenamePresenter(RenameModel model,
            IRefactoringDialogFactory dialogFactory) : base(
            model, dialogFactory)
        { }
        
        public override RenameModel Show()
        {
            return Model.Target == null ? null : base.Show();
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

