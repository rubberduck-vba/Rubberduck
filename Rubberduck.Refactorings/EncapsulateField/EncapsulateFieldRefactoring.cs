using System.Linq;
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
        private readonly EncapsulateFieldRefactoringAction _refactoringAction;
        private readonly EncapsulateFieldPreviewProvider _previewProvider;
        private readonly IEncapsulateFieldModelFactory _modelFactory;

        public EncapsulateFieldRefactoring(
            EncapsulateFieldRefactoringAction refactoringAction,
            EncapsulateFieldPreviewProvider previewProvider,
            IEncapsulateFieldModelFactory encapsulateFieldModelFactory,
            RefactoringUserInteraction<IEncapsulateFieldPresenter, EncapsulateFieldModel> userInteraction,
            ISelectionProvider selectionProvider,
            ISelectedDeclarationProvider selectedDeclarationProvider)
                :base(selectionProvider, userInteraction)
        {
            _refactoringAction = refactoringAction;
            _previewProvider = previewProvider;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _modelFactory = encapsulateFieldModelFactory;
        }

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            var selectedDeclaration = _selectedDeclarationProvider.SelectedDeclaration(targetSelection);

            var isInvalidSelection = selectedDeclaration == null
                || selectedDeclaration.DeclarationType != DeclarationType.Variable
                || selectedDeclaration.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member);

            return isInvalidSelection ? null : selectedDeclaration;
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

            model.StrategyChangedAction = OnStrategyChanged;

            model.ObjectStateFieldChangedAction = OnObjectStateUDTChanged;

            model.ConflictFinder.AssignNoConflictIdentifiers(model.EncapsulationCandidates);

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

        private void OnStrategyChanged(EncapsulateFieldModel model)
        {
            if (model.EncapsulateFieldStrategy == EncapsulateFieldStrategy.UseBackingFields)
            {
                foreach (var objectStateCandidate in model.EncapsulateFieldUseBackingUDTMemberModel.ObjectStateUDTCandidates)
                {
                    objectStateCandidate.IsSelected = !objectStateCandidate.IsExistingDeclaration;
                }
            }

            var candidates = model.EncapsulateFieldStrategy == EncapsulateFieldStrategy.UseBackingFields
                ? model.EncapsulateFieldUseBackingFieldModel.EncapsulationCandidates
                : model.EncapsulateFieldUseBackingUDTMemberModel.EncapsulationCandidates;

            model.ConflictFinder.AssignNoConflictIdentifiers(candidates);
        }

        private void OnObjectStateUDTChanged(EncapsulateFieldModel model)
        {
            model.ConflictFinder.AssignNoConflictIdentifiers(model.EncapsulateFieldUseBackingUDTMemberModel.EncapsulationCandidates);
        }
    }
}
