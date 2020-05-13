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
            return _selectedDeclarationProvider.SelectedDeclaration(targetSelection);
        }

        protected override AnnotateDeclarationModel InitializeModel(Declaration target)
        {
            if (target == null)
            {
                throw new TargetDeclarationIsNullException();
            }

            var targetType = target.DeclarationType;

            if (!targetType.HasFlag(DeclarationType.Member)
                && !targetType.HasFlag(DeclarationType.Module)
                && !targetType.HasFlag(DeclarationType.Variable)
                && !targetType.HasFlag(DeclarationType.Constant))
            {
                throw new InvalidDeclarationTypeException(target);
            }

            return new AnnotateDeclarationModel(target);
        }

        protected override void RefactorImpl(AnnotateDeclarationModel model)
        {
            _annotateDeclarationAction.Refactor(model);
        }
    }
}

