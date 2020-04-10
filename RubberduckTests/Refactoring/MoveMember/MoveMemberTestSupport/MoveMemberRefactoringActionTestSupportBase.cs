using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveMember;
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
            var tdi = new MoveMemberTestsDI(state, rewritingManager);
            return new MoveMemberRefactoringAction(tdi.Resolve<MoveMemberToNewStandardModuleRefactoringAction>(), 
                                                        tdi.Resolve<MoveMemberToExistingStandardModuleRefactoringAction>());
        }

        public static IAddComponentService TestAddComponentService(IProjectsProvider projectsProvider)
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
