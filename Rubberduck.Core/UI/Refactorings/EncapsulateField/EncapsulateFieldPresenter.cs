using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;

namespace Rubberduck.UI.Refactorings.EncapsulateField
{
    internal class EncapsulateFieldPresenter : RefactoringPresenterBase<EncapsulateFieldModel, IRefactoringDialog<EncapsulateFieldModel, IRefactoringView<EncapsulateFieldModel>, IRefactoringViewModel<EncapsulateFieldModel>>, IRefactoringView<EncapsulateFieldModel>, IRefactoringViewModel<EncapsulateFieldModel>>, IEncapsulateFieldPresenter
    {
        public EncapsulateFieldPresenter(EncapsulateFieldModel model,
            IRefactoringDialogFactory dialogFactory) : base(model, dialogFactory) { }

        public override EncapsulateFieldModel Show()
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
