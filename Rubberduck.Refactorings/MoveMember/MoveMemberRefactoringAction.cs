

namespace Rubberduck.Refactorings.MoveMember
{
    public class MoveMemberRefactoringAction : IRefactoringAction<MoveMemberModel>
    {
        private readonly IRefactoringAction<MoveMemberModel> _moveMemberToNewModuleRefactoringAction;
        private readonly IRefactoringAction<MoveMemberModel> _moveMemberToExistingModuleRefactoringAction;

        public MoveMemberRefactoringAction(MoveMemberToNewModuleRefactoringAction moveMemberToNewModuleRefactoring, 
                                            MoveMemberExistingModulesRefactoringAction moveMemberToExistingModuleRefactoring)
        {
            _moveMemberToNewModuleRefactoringAction = moveMemberToNewModuleRefactoring;
            _moveMemberToExistingModuleRefactoringAction = moveMemberToExistingModuleRefactoring;
        }

        public void Refactor(MoveMemberModel model)
        {
            if (model.Destination.IsExistingModule(out _))
            {
                _moveMemberToExistingModuleRefactoringAction.Refactor(model);
                return;
            }

            _moveMemberToNewModuleRefactoringAction.Refactor(model);
        }
    }
}
