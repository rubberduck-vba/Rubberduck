using Rubberduck.Refactorings.AnnotateDeclaration;
using Rubberduck.Refactorings;

namespace Rubberduck.UI.Refactorings.AnnotateDeclaration
{
    internal class AnnotateDeclarationPresenter : RefactoringPresenterBase<AnnotateDeclarationModel>, IAnnotateDeclarationPresenter
    {
        private static readonly DialogData DialogData = DialogData.Create(RefactoringsUI.AnnotateDeclarationDialog_Caption, 500, 400);

        public AnnotateDeclarationPresenter(AnnotateDeclarationModel model, IRefactoringDialogFactory dialogFactory) :
            base(DialogData, model, dialogFactory)
        {}
    }
}