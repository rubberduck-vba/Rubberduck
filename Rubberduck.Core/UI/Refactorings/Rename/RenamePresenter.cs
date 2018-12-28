using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Rename;

namespace Rubberduck.UI.Refactorings.Rename
{
    internal class RenamePresenter : RefactoringPresenterBase<RenameModel, IRefactoringDialog<RenameModel, IRefactoringView<RenameModel>, IRefactoringViewModel<RenameModel>>, IRefactoringView<RenameModel>, IRefactoringViewModel<RenameModel>>, IRenamePresenter
    {
        public RenamePresenter(RenameModel model, IRefactoringDialogFactory dialogFactory) : base(model, dialogFactory) { }

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

            var model = Show();

            if (DialogResult != RefactoringDialogResult.Execute)
            {
                return null;
            }
            
            return model;
        }
    }
}

