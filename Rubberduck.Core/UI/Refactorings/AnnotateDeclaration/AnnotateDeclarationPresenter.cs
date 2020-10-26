using Rubberduck.Refactorings.AnnotateDeclaration;
using Rubberduck.Resources.Annotations;

namespace Rubberduck.UI.Refactorings.AnnotateDeclaration
{
    internal class AnnotateDeclarationPresenter : RefactoringPresenterBase<AnnotateDeclarationModel>, IAnnotateDeclarationPresenter
    {
        private static readonly DialogData DialogData = DialogData.Create(AnnotateDeclarationDialog.Caption, 500, 400);

        public AnnotateDeclarationPresenter(AnnotateDeclarationModel model, IRefactoringDialogFactory dialogFactory) :
            base(DialogData, model, dialogFactory)
        {}
    }
}