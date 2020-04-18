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
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;
        private readonly IMoveMemberModelFactory _modelFactory;

        public MoveMemberRefactoring(
            MoveMemberRefactoringAction refactoringAction,
            RefactoringUserInteraction<IMoveMemberPresenter, MoveMemberModel> userInteraction,
            ISelectionProvider selectionProvider,
            ISelectedDeclarationProvider selectedDeclarationProvider,
            IMoveMemberModelFactory modelFactory)
                : base(selectionProvider, userInteraction)

        {
            _refactoringAction = refactoringAction;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _modelFactory = modelFactory;
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

            return _modelFactory.Create(target);
        }

        protected override void RefactorImpl(MoveMemberModel model)
        {
            _refactoringAction.Refactor(model);
        }
    }
}
