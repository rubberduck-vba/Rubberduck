using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.AnnotateDeclaration
{
    public class AnnotateDeclarationRefactoring : InteractiveRefactoringBase<AnnotateDeclarationModel>
    {
        private readonly IRefactoringAction<AnnotateDeclarationModel> _annotateDeclarationAction;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;

        public AnnotateDeclarationRefactoring(
            AnnotateDeclarationRefactoringAction annotateDeclarationAction,
            ISelectedDeclarationProvider selectedDeclarationProvider,
            ISelectionProvider selectionProvider,
            RefactoringUserInteraction<IAnnotateDeclarationPresenter, AnnotateDeclarationModel> userInteraction)
            : base(selectionProvider, userInteraction)
        {
            _annotateDeclarationAction = annotateDeclarationAction;
            _selectedDeclarationProvider = selectedDeclarationProvider;
        }
        
        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            return _selectedDeclarationProvider.SelectedModule(targetSelection);
        }

        protected override AnnotateDeclarationModel InitializeModel(Declaration target)
        {
            if (target == null)
            {
                throw new TargetDeclarationIsNullException();
            }

            return new AnnotateDeclarationModel(target);
        }

        protected override void RefactorImpl(AnnotateDeclarationModel model)
        {
            _annotateDeclarationAction.Refactor(model);
        }
    }
}

