using Rubberduck.Refactorings.EncapsulateField;

namespace Rubberduck.UI.Refactorings.EncapsulateField
{
    internal class EncapsulateFieldPresenter : RefactoringPresenterBase<EncapsulateFieldModel>, IEncapsulateFieldPresenter
    {
        private static readonly DialogData DialogData =
            DialogData.Create(Rubberduck.Resources.Refactorings.EncapsulateField.Caption, 800, 900);

        public EncapsulateFieldPresenter(EncapsulateFieldModel model,
            IRefactoringDialogFactory dialogFactory) : base(DialogData, model, dialogFactory) { }
    }
}
