using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.MoveMember
{
    public class MoveMemberRefactoring : InteractiveRefactoringBase<MoveMemberModel>
    {
        private readonly MoveMemberRefactoringAction _refactoringAction;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;
        private readonly IMoveMemberStrategyFactory _strategyFactory;
        private readonly IMoveMemberEndpointFactory _endpointFactory;
        private readonly IMoveMemberRefactoringPreviewerFactory _previewerFactory;

        public MoveMemberRefactoring(
            MoveMemberRefactoringAction refactoringAction,
            RefactoringUserInteraction<IMoveMemberPresenter, MoveMemberModel> userInteraction,
            IDeclarationFinderProvider declarationFinderProvider,
            ISelectionProvider selectionProvider,
            ISelectedDeclarationProvider selectedDeclarationProvider,
            IMoveMemberStrategyFactory strategyFactory,
            IMoveMemberEndpointFactory enpointFactory,
            IMoveMemberRefactoringPreviewerFactory previewerFactory)
                : base(selectionProvider, userInteraction)

        {
            _refactoringAction = refactoringAction;
            _declarationFinderProvider = declarationFinderProvider;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _strategyFactory = strategyFactory;
            _previewerFactory = previewerFactory;
            _endpointFactory = enpointFactory;
        }

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            var selected = _selectedDeclarationProvider.SelectedDeclaration(targetSelection);
            if (selected.IsMember()
                || selected.IsModuleConstant()
                || (selected.IsMemberVariable() && !selected.HasPrivateAccessibility()))
            {
                return selected;
            }

            return null;
        }

        protected override MoveMemberModel InitializeModel(Declaration target)
        {
            if (target == null) { throw new TargetDeclarationIsNullException(); }

            return new MoveMemberModel(target,
                                        _declarationFinderProvider,
                                        _strategyFactory,
                                        _endpointFactory);
        }

        protected override void RefactorImpl(MoveMemberModel model)
        {
            _refactoringAction.Refactor(model);
        }
    }
}
