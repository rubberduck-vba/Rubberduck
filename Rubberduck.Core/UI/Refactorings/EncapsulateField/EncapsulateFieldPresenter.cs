using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Refactorings;

namespace Rubberduck.UI.Refactorings.EncapsulateField
{
    internal class EncapsulateFieldPresenter : RefactoringPresenterBase<EncapsulateFieldModel>, IEncapsulateFieldPresenter
    {
        private static readonly DialogData DialogData =
            DialogData.Create(RefactoringsUI.EncapsulateField_Caption, 800, 900);

        public EncapsulateFieldPresenter(EncapsulateFieldModel model,
            IRefactoringDialogFactory dialogFactory) : base(DialogData, model, dialogFactory) { }
    }
}
