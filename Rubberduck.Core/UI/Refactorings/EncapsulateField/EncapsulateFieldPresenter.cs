using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Resources;

namespace Rubberduck.UI.Refactorings.EncapsulateField
{
    internal class EncapsulateFieldPresenter : RefactoringPresenterBase<EncapsulateFieldModel>, IEncapsulateFieldPresenter
    {
        private static readonly DialogData DialogData =
            DialogData.Create(RubberduckUI.EncapsulateField_Caption, 305, 667);

        public EncapsulateFieldPresenter(EncapsulateFieldModel model,
            IRefactoringDialogFactory dialogFactory) : base(DialogData, model, dialogFactory) { }

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
