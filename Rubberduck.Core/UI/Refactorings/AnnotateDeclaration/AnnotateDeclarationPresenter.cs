using Rubberduck.Refactorings.AnnotateDeclaration;
using Rubberduck.Resources;

namespace Rubberduck.UI.Refactorings.AnnotateDeclaration
{
    internal class AnnotateDeclarationPresenter : RefactoringPresenterBase<AnnotateDeclarationModel>, IAnnotateDeclarationPresenter
    {
        private static readonly DialogData DialogData = DialogData.Create(RubberduckUI.AnnotateDeclarationDialog_Caption, 400, 300);

        public AnnotateDeclarationPresenter(AnnotateDeclarationModel model, IRefactoringDialogFactory dialogFactory) :
            base(DialogData, model, dialogFactory)
        {}
    }
}