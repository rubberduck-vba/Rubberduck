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
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;
        private readonly IRewritingManager _rewritingManager;
        private readonly EncapsulateFieldRefactoringAction _refactoringAction;
        private readonly EncapsulateFieldPreviewProvider _previewProvider;

        public EncapsulateFieldRefactoring(
                EncapsulateFieldRefactoringAction refactoringAction,
                EncapsulateFieldPreviewProvider previewProvider,
                IDeclarationFinderProvider declarationFinderProvider,
                RefactoringUserInteraction<IEncapsulateFieldPresenter, EncapsulateFieldModel> userInteraction,
                IRewritingManager rewritingManager,
                ISelectionProvider selectionProvider,
                ISelectedDeclarationProvider selectedDeclarationProvider)
            :base(selectionProvider, userInteraction)
        {
            _refactoringAction = refactoringAction;
            _previewProvider = previewProvider;
            _declarationFinderProvider = declarationFinderProvider;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _rewritingManager = rewritingManager;
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

            var builder = new EncapsulateFieldElementsBuilder(_declarationFinderProvider, target.QualifiedModuleName);

            var selected = builder.Candidates.Single(c => c.Declaration == target);
            selected.EncapsulateFlag = true;

            var model = new EncapsulateFieldModel(
                                target,
                                builder.Candidates,
                                builder.ObjectStateUDTCandidates,
                                builder.DefaultObjectStateUDT,
                                _declarationFinderProvider,
                                builder.ValidationsProvider)
            {
                PreviewProvider = _previewProvider,
                ObjectStateUDTField = builder.ObjectStateUDT,
                EncapsulateFieldStrategy = builder.ObjectStateUDT != null ? EncapsulateFieldStrategy.ConvertFieldsToUDTMembers : EncapsulateFieldStrategy.UseBackingFields,
            };

            return model;
        }

        private EncapsulateFieldStrategy ApplyStrategy(IObjectStateUDT objStateUDT)
        {
            return objStateUDT != null ? EncapsulateFieldStrategy.ConvertFieldsToUDTMembers : EncapsulateFieldStrategy.UseBackingFields;
        }

        protected override void RefactorImpl(EncapsulateFieldModel model)
        {
            if (!model.SelectedFieldCandidates.Any()) { return; }

            _refactoringAction.Refactor(model);
        }
    }
}
