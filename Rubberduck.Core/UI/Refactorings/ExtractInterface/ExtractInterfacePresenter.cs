using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ExtractInterface;

namespace Rubberduck.UI.Refactorings.ExtractInterface
{
    internal class ExtractInterfacePresenter : RefactoringPresenterBase<ExtractInterfaceModel, IRefactoringDialog<ExtractInterfaceModel, IRefactoringView<ExtractInterfaceModel>, IRefactoringViewModel<ExtractInterfaceModel>>, IRefactoringView<ExtractInterfaceModel>, IRefactoringViewModel<ExtractInterfaceModel>>, IExtractInterfacePresenter
    {
        public ExtractInterfacePresenter(ExtractInterfaceModel model,
            IRefactoringDialogFactory dialogFactory) : base(model, dialogFactory) { }

        public override ExtractInterfaceModel Show()
        {
            if (Model.TargetDeclaration == null)
            {
                return null;
            }

            var model = base.Show();
            return DialogResult != RefactoringDialogResult.Execute ? null : model;
        }
    }
}
