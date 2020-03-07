﻿using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

        public static IEnumerable<IMoveMemberRefactoringStrategy> RetrieveStrategies(IDeclarationFinderProvider declarationFinderProvider, string declarationName, DeclarationType declarationType, IRewritingManager rewritingManager)
        {
            var target = declarationFinderProvider.DeclarationFinder.DeclarationsWithType(declarationType)
                 .Single(declaration => declaration.IdentifierName == declarationName);

            var model = new MoveMemberModel(target, declarationFinderProvider);

            model.ChangeDestination(DEFAULT_DESTINATION_MODULE_NAME);

            if (MoveMemberObjectsFactory.TryCreateStrategy(model, out var strategy))
            {
                return new IMoveMemberRefactoringStrategy[] { strategy };
            }

            return Enumerable.Empty<IMoveMemberRefactoringStrategy>(); ;
        }

        public static MoveMemberModel CreateModelAndDefineMove(IVBE vbe, TestMoveDefinition moveDefinition, RubberduckParserState state, IRewritingManager rewritingManager)
        {
            var sourceModule = state.DeclarationFinder.ModuleDeclaration(GetQMN(vbe, moveDefinition.SourceModuleName));
            var member = state.DeclarationFinder.Members(sourceModule).FirstOrDefault(m => m.IdentifierName.Equals(moveDefinition.SelectedElement));
            var destinationModule = state.DeclarationFinder.ModuleDeclaration(GetQMN(vbe, moveDefinition.DestinationModuleName));
            var model = new MoveMemberModel(member, state);

            model.ChangeDestination(destinationModule);
            return model;
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

        public static QualifiedModuleName GetQMN(IVBE VBE, string moduleName, string projectName = DEFAULT_PROJECT_NAME)
        {
            var project = VBE.VBProjects.Single(item => item.Name == projectName);
            var component = project.VBComponents.SingleOrDefault(c => c.Name == moduleName);
            using (component)
            {
                return component != null ? new QualifiedModuleName(component) : new QualifiedModuleName(project);
            }
        }

        public static bool OccursOnce(string toFind, string content)
        {
            var firstIdx = content.IndexOf(toFind);
            var lastIdx = content.LastIndexOf(toFind);
            return firstIdx == lastIdx;
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