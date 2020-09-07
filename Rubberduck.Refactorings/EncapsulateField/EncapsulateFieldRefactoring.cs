using System.Linq;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public enum EncapsulateFieldStrategy
    {
        UseBackingFields,
        ConvertFieldsToUDTMembers
    }

    public class EncapsulateFieldRefactoring : InteractiveRefactoringBase<EncapsulateFieldModel>
    {
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;
        private readonly IRewritingManager _rewritingManager;
        private readonly EncapsulateFieldRefactoringAction _refactoringAction;
        private readonly EncapsulateFieldPreviewProvider _previewProvider;
        private readonly IEncapsulateFieldModelFactory _modelFactory;

        public EncapsulateFieldRefactoring(
            EncapsulateFieldRefactoringAction refactoringAction,
            EncapsulateFieldPreviewProvider previewProvider,
            IEncapsulateFieldModelFactory encapsulateFieldModelFactory,
            RefactoringUserInteraction<IEncapsulateFieldPresenter, EncapsulateFieldModel> userInteraction,
            IRewritingManager rewritingManager,
            ISelectionProvider selectionProvider,
            ISelectedDeclarationProvider selectedDeclarationProvider)
                :base(selectionProvider, userInteraction)
        {
            _refactoringAction = refactoringAction;
            _previewProvider = previewProvider;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _rewritingManager = rewritingManager;
            _modelFactory = encapsulateFieldModelFactory;
        }

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            var selectedDeclaration = _selectedDeclarationProvider.SelectedDeclaration(targetSelection);
            if (selectedDeclaration == null
                || selectedDeclaration.DeclarationType != DeclarationType.Variable
                || selectedDeclaration.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member))
            {
                return null;
            }

            return selectedDeclaration;
        }

        protected override EncapsulateFieldModel InitializeModel(Declaration target)
        {
            if (target == null)
            {
                throw new TargetDeclarationIsNullException();
            }

            if (!target.DeclarationType.Equals(DeclarationType.Variable))
            {
                throw new InvalidDeclarationTypeException(target);
            }

            var model = _modelFactory.Create(target);

            model.PreviewProvider = _previewProvider;

            return model;
        }

        protected override void RefactorImpl(EncapsulateFieldModel model)
        {
            if (!model.SelectedFieldCandidates.Any())
            {
                return;
            }

            _refactoringAction.Refactor(model);
        }
    }
}
