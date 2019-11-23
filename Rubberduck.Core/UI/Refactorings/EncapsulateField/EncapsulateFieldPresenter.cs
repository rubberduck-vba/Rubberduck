using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Resources;

namespace Rubberduck.UI.Refactorings.EncapsulateField
{
    internal class EncapsulateFieldPresenter : RefactoringPresenterBase<EncapsulateFieldModel>, IEncapsulateFieldPresenter
    {
        private static readonly DialogData DialogData =
            DialogData.Create(RubberduckUI.EncapsulateField_Caption, 800, 800);

        public EncapsulateFieldPresenter(EncapsulateFieldModel model,
            IRefactoringDialogFactory dialogFactory) : base(DialogData, model, dialogFactory) { }

        public override EncapsulateFieldModel Show()
        {
            return Model.TargetDeclaration == null ? null : base.Show();
        }
    }
}
