using Moq;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.Refactorings.Rename;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RubberduckTests.Refactoring.MoveMember
{
    public class MoveMemberTestSupport
    {
        public const string DEFAULT_PROJECT_NAME = MockVbeBuilder.TestProjectName;
        public const string DEFAULT_SOURCE_MODULE_NAME = "DfltSrcStd";
        public const string DEFAULT_SOURCE_CLASS_NAME = "DfltSrcClass";
        public const string DEFAULT_SOURCE_FORM_NAME = "DfltSrcForm";
        public const string DEFAULT_DESTINATION_MODULE_NAME = "DfltDestStd";
        public const string DEFAULT_DESTINATION_CLASS_NAME = "DfltDestClass";

        public static T ParseAndTest<T>(IVBE vbe, Func<RubberduckParserState, T> testFunc)
        {
            T result = default;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                result = testFunc(state);
            }
            return result;
        }

        public static T ParseAndTest<T>(IVBE vbe, Func<RubberduckParserState, IVBE, IRewritingManager, T> testFunc)
        {
            T result = default;
            (RubberduckParserState state, IRewritingManager rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                result = testFunc(state, vbe, rewritingManager);
            }
            return result;
        }

        public static T ParseAndTest<T>(Func<RubberduckParserState, IVBE, IRewritingManager, T> testFunc, params (string moduleName, string content, ComponentType componentType)[] modules)
        {
            T result = default;
            var vbe = MockVbeBuilder.BuildFromModules(modules).Object;
            (RubberduckParserState state, IRewritingManager rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                result = testFunc(state, vbe, rewritingManager);
            }
            return result;
        }

        public static MoveMemberRefactoring CreateRefactoring(IRewritingManager rewritingManager, RubberduckParserState state, RefactoringUserInteraction<IMoveMemberPresenter, MoveMemberModel> userInteraction, ISelectionService selectionService)
        {
            var selectedDeclarationService = new SelectedDeclarationProvider(selectionService, state);
            var tdi = new MoveMemberTestsDI(state, rewritingManager);
            return new MoveMemberRefactoring(tdi.Resolve<MoveMemberRefactoringAction>(),
                                                userInteraction, 
                                                state, 
                                                selectionService, 
                                                selectedDeclarationService,
                                                tdi.Resolve<MoveMemberStrategyFactory>(),
                                                tdi.Resolve<MoveMemberEndpointFactory>(),
                                                tdi.Resolve<MoveMemberRefactoringPreviewerFactory>()
                                                );
        }

        public static MoveMemberModel CreateRefactoringModel(Declaration target, RubberduckParserState state, IRewritingManager rewritingManager)
        {
            var tdi = new MoveMemberTestsDI(state, rewritingManager);
            return new MoveMemberModel(target, 
                                        state,
                                        tdi.Resolve<MoveMemberStrategyFactory>(),
                                        tdi.Resolve<MoveMemberEndpointFactory>()
                                        );
        }

        public static IEnumerable<IMoveMemberRefactoringStrategy> RetrieveStrategies(RubberduckParserState state, string declarationName, DeclarationType declarationType, IRewritingManager rewritingManager)
        {
            var target = state.DeclarationFinder.DeclarationsWithType(declarationType)
                 .Single(declaration => declaration.IdentifierName == declarationName);

            var model = CreateRefactoringModel(target, state, rewritingManager);

            model.ChangeDestination(DEFAULT_DESTINATION_MODULE_NAME);

            if (model.TryFindApplicableStrategy(out var strategy))
            {
                return new IMoveMemberRefactoringStrategy[] { strategy };
            }

            return Enumerable.Empty<IMoveMemberRefactoringStrategy>(); ;
        }

        public static IVBE BuildVBEStub(TestMoveDefinition moveDefinition, string sourceContent)
        {
            if (moveDefinition.CreateNewModule)
            {
                moveDefinition.SetEndpointContent(sourceContent);
                return MockVbeBuilder.BuildFromModules(moveDefinition.ModuleDefinitions.Select(tc => tc.AsTuple)).Object;
            }
            moveDefinition.SetEndpointContent(sourceContent, null);
            return MockVbeBuilder.BuildFromModules(moveDefinition.ModuleDefinitions.Select(tc => tc.AsTuple)).Object;
        }

        public static bool OccursOnce(string toFind, string content)
        {
            var firstIdx = content.IndexOf(toFind);
            var lastIdx = content.LastIndexOf(toFind);
            return firstIdx == lastIdx && firstIdx > -1;
        }

        public static (string moduleName, string content, ComponentType componentType) EndpointToSourceTuple(MoveEndpoints endpoints, string content)
        {
            switch (endpoints)
            {
                case MoveEndpoints.FormToStd:
                    return (DEFAULT_SOURCE_FORM_NAME, content, ComponentType.UserForm);
                case MoveEndpoints.ClassToStd:
                    return (DEFAULT_SOURCE_CLASS_NAME, content, ComponentType.ClassModule);
                case MoveEndpoints.StdToStd:
                    return (DEFAULT_SOURCE_MODULE_NAME, content, ComponentType.StandardModule);
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        public static (string moduleName, string content, ComponentType componentType) EndpointToDestinationTuple(MoveEndpoints endpoints, string content)
        {
            switch (endpoints)
            {
                case MoveEndpoints.FormToStd:
                case MoveEndpoints.ClassToStd:
                case MoveEndpoints.StdToStd:
                    return (DEFAULT_DESTINATION_MODULE_NAME, content, ComponentType.StandardModule);
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        public static string ClassInstantiationBoilerPlate(string instanceIdentifier, string className)
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
