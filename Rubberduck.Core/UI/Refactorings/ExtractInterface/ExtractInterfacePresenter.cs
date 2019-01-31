using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.Resources;

namespace Rubberduck.UI.Refactorings.ExtractInterface
{
    internal class ExtractInterfacePresenter : RefactoringPresenterBase<ExtractInterfaceModel>, IExtractInterfacePresenter
    {
        private static readonly DialogData DialogData = DialogData.Create(RubberduckUI.ExtractInterface_Caption, 339, 459);

        public ExtractInterfacePresenter(ExtractInterfaceModel model,
            IRefactoringDialogFactory dialogFactory) : base(DialogData, model, dialogFactory) { }

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
