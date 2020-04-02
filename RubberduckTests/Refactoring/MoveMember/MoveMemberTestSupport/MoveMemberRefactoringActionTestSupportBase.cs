using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.Refactorings.Rename;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SourceCodeHandling;
using Rubberduck.VBEditor.Utility;
using System.Linq;

namespace RubberduckTests.Refactoring.MoveMember
{
    public class MoveMemberRefactoringActionTestSupportBase : RefactoringActionTestBase<MoveMemberModel>
    {
        protected override IRefactoringAction<MoveMemberModel> TestBaseRefactoring(RubberduckParserState state, IRewritingManager rewritingManager)
        {
            var addComponentService = TestAddComponentService(state?.ProjectsProvider);
            var renameAction = new RenameCodeDefinedIdentifierRefactoringAction(state, state?.ProjectsProvider, rewritingManager);
            var existingDestinationModuleRefactoring = new MoveMemberExistingModulesRefactoringAction(renameAction, state, rewritingManager);
            var newDestinationModuleRefactoring = new MoveMemberToNewModuleRefactoringAction(existingDestinationModuleRefactoring, renameAction, state, rewritingManager, addComponentService);
            return new MoveMemberRefactoringAction(newDestinationModuleRefactoring, existingDestinationModuleRefactoring);
        }

        private static IAddComponentService TestAddComponentService(IProjectsProvider projectsProvider)
        {
            var sourceCodeHandler = new CodeModuleComponentSourceCodeHandler();
            return new AddComponentService(projectsProvider, sourceCodeHandler, sourceCodeHandler);
        }

        protected MoveMemberRefactorResults ExecuteTest(TestMoveDefinition moveDefinition)
        {
            var results = RefactoredCode(moveDefinition.ModelBuilder, moveDefinition.ModuleTuples.ToArray());
            return new MoveMemberRefactorResults(moveDefinition, results, moveDefinition.StrategyName);
        }
    }
}
