using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Rename;

namespace Rubberduck.UI.Refactorings.Rename
{
    internal class RenamePresenter : RefactoringPresenterBase<RenameModel, IRefactoringDialog<RenameModel, IRefactoringView<RenameModel>, IRefactoringViewModel<RenameModel>>, IRefactoringView<RenameModel>, IRefactoringViewModel<RenameModel>>, IRenamePresenter
    {
        private readonly RenameViewModel _viewModel;

        public RenamePresenter(RenameModel model, IRefactoringDialogFactory dialogFactory) : base(model, dialogFactory)
        {
            _viewModel = dialogFactory.CreateViewModel<RenameModel, RenameViewModel>(model);
        }

        public override IRefactoringViewModel<RenameModel> ViewModel => _viewModel;

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
            _viewModel.Target = target;

            var model = Show();

            if (DialogResult != RefactoringDialogResult.Execute)
            {
                return null;
            }

            //Model.NewName = _viewModel.NewName;
            return model;
        }
    }
}

