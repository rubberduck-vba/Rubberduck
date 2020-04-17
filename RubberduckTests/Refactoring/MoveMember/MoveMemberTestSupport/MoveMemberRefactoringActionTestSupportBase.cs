using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System;
using System.Linq;

namespace RubberduckTests.Refactoring.MoveMember
{
    public class MoveMemberRefactoringActionTestSupportBase : RefactoringActionTestBase<MoveMemberModel>
    {
        protected override IRefactoringAction<MoveMemberModel> TestBaseRefactoring(RubberduckParserState state, IRewritingManager rewritingManager)
        {
            MoveMemberTestsResolver serviceLocator = new MoveMemberTestsResolver(state, rewritingManager);
            return serviceLocator.Resolve<MoveMemberRefactoringAction>();
        }

        protected MoveMemberRefactorResults RefactorTargets((string ID, DeclarationType DecType) memberToMove, MoveEndpoints endpoints, string sourceContent, string destinationContent, Func<MoveMemberModel, MoveMemberModel> modelAdjustment)
        {
            var sourceTuple = endpoints.ToSourceTuple(sourceContent);
            var destinationTuple = endpoints.ToDestinationTuple(destinationContent);

            return RefactorTargets(memberToMove, endpoints, modelAdjustment, sourceTuple, destinationTuple);
        }

        protected MoveMemberRefactorResults RefactorTargets((string ID, DeclarationType DecType) memberToMove, MoveEndpoints endpoints, Func<MoveMemberModel, MoveMemberModel> modelAdjustment, params (string, string, ComponentType)[] modules)
        {
            return ExecuteRefactoring(memberToMove, endpoints, modelAdjustment, modules);
        }

        protected MoveMemberRefactorResults RefactorSingleTarget((string ID, DeclarationType DecType) memberToMove, MoveEndpoints endpoints, string sourceContent, string destinationContent = null)
        {
            return RefactorSingleTarget(memberToMove, endpoints, endpoints.ToModulesTuples(sourceContent, destinationContent ?? string.Empty));
        }

        //Takes a sourceModuleName parameter for cases where another module has an identical declaration name and type. (e.g., name collision tests)
        protected MoveMemberRefactorResults RefactorSingleTarget((string ID, DeclarationType DecType) memberToMove, MoveEndpoints endpoints, string sourceModuleName, string sourceContent, string destinationContent = null)
        {
            return RefactorSingleTarget(memberToMove, endpoints, sourceModuleName, endpoints.ToModulesTuples(sourceContent, destinationContent ?? string.Empty));
        }

        protected MoveMemberRefactorResults RefactorSingleTarget((string ID, DeclarationType DecType) memberToMove, MoveEndpoints endpoints, params (string, string, ComponentType)[] moduleTuples)
        {
            MoveMemberModel modelAdjustment(MoveMemberModel model)
            {
                model.ChangeDestination(endpoints.DestinationModuleName(), ComponentType.StandardModule);
                return model;
            }

            var vbe = MockVbeBuilder.BuildFromModules(moduleTuples);
            var results = RefactoredCode(vbe.Object, state => TestModel(state, modelAdjustment, memberToMove.ID, memberToMove.DecType));

            return new MoveMemberRefactorResults(endpoints, results);
        }

        //Takes a sourceModuleName parameter for cases where another module has an identical declaration name and type. (e.g., name collision tests)
        protected MoveMemberRefactorResults RefactorSingleTarget((string ID, DeclarationType DecType) memberToMove, MoveEndpoints endpoints, string sourceModuleName, params (string, string, ComponentType)[] moduleTuples)
        {
            MoveMemberModel modelAdjustment(MoveMemberModel model)
            {
                model.ChangeDestination(endpoints.DestinationModuleName(), ComponentType.StandardModule);
                return model;
            }

            var vbe = MockVbeBuilder.BuildFromModules(moduleTuples);
            var results = RefactoredCode(vbe.Object, state => TestModel(state, modelAdjustment, memberToMove.ID, memberToMove.DecType, sourceModuleName));

            return new MoveMemberRefactorResults(endpoints, results);
        }

        protected void ExecuteSingleTargetMoveThrowsExceptionTest((string ID, DeclarationType DecType) memberToMove, MoveEndpoints endpoints, string sourceContent, string destinationContent = null)
        {
            ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints, endpoints.ToModulesTuples(sourceContent, destinationContent ?? string.Empty));
        }

        protected void ExecuteSingleTargetMoveThrowsExceptionTest((string ID, DeclarationType DecType) memberToMove, MoveEndpoints endpoints, params (string, string, ComponentType)[] moduleTuples)
        {
            MoveMemberModel modelAdjustment(MoveMemberModel model)
            {
                model.ChangeDestination(endpoints.DestinationModuleName(), ComponentType.StandardModule);
                return model;
            }

            Assert.Throws<MoveMemberUnsupportedMoveException>(() => ExecuteRefactoring(memberToMove, endpoints, modelAdjustment, moduleTuples));
        }

        protected MoveMemberRefactorResults ExecuteRefactoring((string ID, DeclarationType DecType) memberToMove, MoveEndpoints endpoints, Func<MoveMemberModel, MoveMemberModel> modelAdjustment, params (string, string, ComponentType)[] moduleTuples)
        {
            var vbe = MockVbeBuilder.BuildFromModules(moduleTuples);
            var results = RefactoredCode(vbe.Object, state => TestModel(state, modelAdjustment, memberToMove.ID, memberToMove.DecType));

            return new MoveMemberRefactorResults(endpoints, results);
        }

        private static MoveMemberModel TestModel(RubberduckParserState state, Func<MoveMemberModel, MoveMemberModel> modelAdjustment, string targetID, DeclarationType declarationType)
        {
            var target = state.DeclarationFinder.DeclarationsWithType(declarationType)
                .Where(d => d.IdentifierName == targetID).Single();
            var model = MoveMemberTestsResolver.CreateRefactoringModel(target, state);

            return modelAdjustment(model);
        }

        //Takes a sourceModuleName parameter for cases where another module has an identical declaration name and type. (e.g., name collision tests)
        private static MoveMemberModel TestModel(RubberduckParserState state, Func<MoveMemberModel, MoveMemberModel> modelAdjustment, string targetID, DeclarationType declarationType, string sourceModuleName)
        {
            var target = state.DeclarationFinder.DeclarationsWithType(declarationType)
                .Where(d => d.IdentifierName == targetID && d.QualifiedModuleName.ComponentName == sourceModuleName).Single();
            var model = MoveMemberTestsResolver.CreateRefactoringModel(target, state);

            return modelAdjustment(model);
        }

        protected RubberduckParserState CreateAndParse(MoveEndpoints endpoints, string sourceContent, string destinationContent)
        {
            var modules = endpoints.ToModulesTuples(sourceContent, destinationContent);
            var vbe = MockVbeBuilder.BuildFromModules(modules).Object;
            return MockParser.CreateAndParse(vbe);
        }

        protected string ClassInstantiationBoilerPlate(string instanceIdentifier, string className)
        {
            var declaration = $"Private {instanceIdentifier} As {className}";
            var instantiation =
    $@"
Public Sub Initialize()
    Set {instanceIdentifier} = new {className}
End Sub
";
            return $"{declaration}{Environment.NewLine}{instantiation}";
        }
    }
}
