using Moq;
using NUnit.Framework;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;
using Support = RubberduckTests.Refactoring.MoveMember.MoveMemberTestSupport;

namespace RubberduckTests.Refactoring.MoveMember
{
    public struct MoveMemberRefactorResults
    {
        private readonly IDictionary<string, string> _results;
        private readonly string _sourceModuleName;
        private readonly string _destinationModuleName;
        private readonly string _strategyName;

        public MoveMemberRefactorResults(TestMoveDefinition moveDefinition, IDictionary<string, string> refactorResults, string strategy = null)
        {
            _results = refactorResults;
            _sourceModuleName = moveDefinition.SourceModuleName;
            _destinationModuleName = moveDefinition.DestinationModuleName;
            _strategyName = strategy;
        }

        public string this[string moduleName]
        {
            get => _results[moduleName];
        }

        public string Source => _results[_sourceModuleName];
        public string Destination => _results[_destinationModuleName];
        public string StrategyName => _strategyName;
    }

    public class MoveMemberTestsBase : InteractiveRefactoringTestBase<IMoveMemberPresenter, MoveMemberModel>
    {
        protected override IRefactoring TestRefactoring(IRewritingManager rewritingManager, RubberduckParserState state, IRefactoringPresenterFactory factory, ISelectionService selectionService)
        {
            var uiDispatcherMock = new Mock<IUiDispatcher>();
            uiDispatcherMock
                .Setup(m => m.Invoke(It.IsAny<Action>()))
                .Callback((Action action) => action.Invoke());
            return new MoveMemberRefactoring(state, state, null, factory, rewritingManager, selectionService, new SelectedDeclarationProvider(selectionService, state), uiDispatcherMock.Object);
        }

        protected MoveMemberRefactorResults RefactoredCode(TestMoveDefinition moveDefinition, string sourceContent, string destinationContent = null, Type expectedException = null, bool executeViaActiveSelection = false, params (string identifier, DeclarationType declarationType)[] additionalElements)
        {
            string strategyName = null;
            MoveMemberModel PresenterAdjustment(MoveMemberModel model)
            {
                model.ChangeDestination(moveDefinition.DestinationModuleName);
                var moveCandidates = model.MoveableMembers;
                foreach (var element in additionalElements)
                {
                    var target = model.DeclarationFinderProvider.DeclarationFinder.DeclarationsWithType(element.declarationType)
                        .Where(t => t.ParentDeclaration.IdentifierName == moveDefinition.SourceModuleName)
                        .Single(declaration => declaration.IdentifierName == element.Item1);

                    var moveable = moveCandidates.Single(mv => mv.IdentifierName == element.identifier);
                    moveable.IsSelected = true;
                }

                strategyName = GetStrategyName(model);
                return model;
            }


            var vbeStub = Support.BuildVBEStub(moveDefinition, sourceContent, destinationContent);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbeStub);
            using (state)
            {
                var target = state.DeclarationFinder.DeclarationsWithType(moveDefinition.SelectedDeclarationType)
                    .Where(t => t.ParentDeclaration.IdentifierName == moveDefinition.SourceModuleName)
                    .Single(declaration => declaration.IdentifierName == moveDefinition.SelectedElement);

                var refactoring = TestRefactoring(rewritingManager, state, PresenterAdjustment);

                if (expectedException != null)
                {
                    Assert.Throws(expectedException, () => refactoring.Refactor(target));
                }
                else
                {
                    refactoring.Refactor(target);
                }

                var results = vbeStub.ActiveVBProject.VBComponents
                    .ToDictionary(component => component.Name, component => component.CodeModule.Content());
                return new MoveMemberRefactorResults(moveDefinition, results, strategyName);
            }
        }

        private string GetStrategyName(MoveMemberModel model)
        {
            if (MoveMemberObjectsFactory.TryCreateStrategy(model, out var strategy))
            {
                return strategy.GetType().Name;
            }
            return null;
        }

        protected MoveMemberRefactorResults RefactoredCode_UserSetsDestinationModuleName(TestMoveDefinition moveDefinition, string sourceContent)
        {
            string strategyName = null;
            MoveMemberModel PresenterAdjustment(MoveMemberModel model)
            {
                model.ChangeDestination(moveDefinition.DestinationModuleName);
                strategyName = GetStrategyName(model);
                return model;
            }


            var vbeStub = Support.BuildVBEStub(moveDefinition, sourceContent);
            var results = RefactoredCode(vbeStub, moveDefinition.SelectedElement, moveDefinition.SelectedDeclarationType, PresenterAdjustment);
            return new MoveMemberRefactorResults(moveDefinition, results, strategyName);
        }

        protected string RetrievePreviewAfterUserInput(TestMoveDefinition moveDefinition, string sourceContent, (string declarationName, DeclarationType declarationType) memberToMove)
        {
            MoveMemberModel PresenterAdjustment(MoveMemberModel model)
            {
                model.ChangeDestination(moveDefinition.DestinationModuleName);
                return model;
            }

            var vbe = Support.BuildVBEStub(moveDefinition, sourceContent);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var target = state.DeclarationFinder.DeclarationsWithType(memberToMove.declarationType)
                                    .Single(declaration => declaration.IdentifierName == memberToMove.declarationName);

                var refactoring = TestRefactoring(rewritingManager, state, PresenterAdjustment);
                if (refactoring is IMoveMemberRefactoringTestAccess testAccess)
                {
                    var model = testAccess.TestUserInteractionOnly(target, PresenterAdjustment);
                    return testAccess.PreviewModuleContent(model, PreviewModule.Destination);
                }
                throw new InvalidCastException();
            }
        }
    }

    public class MoveMemberTestSupport
    {
        public const string DEFAULT_PROJECT_NAME = MockVbeBuilder.TestProjectName;
        public const string DEFAULT_SOURCE_MODULE_NAME = "DfltSrcStd";
        public const string DEFAULT_SOURCE_CLASS_NAME = "DfltSrcClass";
        public const string DEFAULT_SOURCE_FORM_NAME = "DfltSrcForm";
        public const string DEFAULT_DESTINATION_MODULE_NAME = "DfltDestStd";
        public const string DEFAULT_DESTINATION_CLASS_NAME = "DfltDestClass";

        public static string PARAM_PREFIX = MoveMemberResources.Prefix_Parameter;

        //If destinationOriginalContent is null, the refactoring is to an existing empty module
        public static MoveMemberTestResult RefactorToExistingDestinationModule(TestMoveDefinition moveDefinition, string sourceOriginalContent, string destinationOriginalContent = null)
        {
            var results = new MoveMemberTestResult(moveDefinition);

            void ThisTest(RubberduckParserState state, IVBE vbe, IRewritingManager rewritingManager)
            {
                ExecuteMoveMemberRefactoring(vbe, moveDefinition, state, rewritingManager);
            }

            var vbeStub = BuildVBEStub(moveDefinition, sourceOriginalContent, destinationOriginalContent);
            ParseAndTest(vbeStub, ThisTest);

            foreach (var moduleDefinition in moveDefinition.ModuleDefinitions)
            {
                results.Add(moduleDefinition.ModuleName, RetrieveModuleContent(vbeStub, moduleDefinition.ModuleName));
            }
            return results;
        }

        public static void ExecuteMoveMemberRefactoring(IVBE vbe, TestMoveDefinition moveDefinition, RubberduckParserState state, IRewritingManager rewritingManager, IMessageBox msgBox = null)
        {
            var member = state.DeclarationFinder.AllUserDeclarations.FirstOrDefault(d => d.IdentifierName.Equals(moveDefinition.SelectedElement));
            var destinationModule = state.DeclarationFinder.ModuleDeclaration(GetQMN(vbe, moveDefinition.DestinationModuleName));

            var model = new MoveMemberModel(member, state, null, new MoveMemberObjectsFactory(state));
            model.ChangeDestination(destinationModule);

            var selectionService = MockedSelectionService(vbe.GetActiveSelection());
            if (msgBox == null)
            {
                msgBox = new Mock<IMessageBox>().Object;
            }

            var presenterFactoryStub = CreatePresenterFactoryStub(model);

            var uiDispatcherMock = new Mock<IUiDispatcher>();
            uiDispatcherMock
                .Setup(m => m.Invoke(It.IsAny<Action>()))
                .Callback((Action action) => action.Invoke());
            var moveMemberRefactoring = new MoveMemberRefactoring(state, state, msgBox, presenterFactoryStub.Object, rewritingManager, selectionService, new SelectedDeclarationProvider(selectionService, state), uiDispatcherMock.Object);

            moveMemberRefactoring.Refactor();
        }

        public static ISelectionService MockedSelectionService(QualifiedSelection? initialSelection)
        {
            QualifiedSelection? activeSelection = initialSelection;
            var selectionServiceMock = new Mock<ISelectionService>();
            selectionServiceMock.Setup(m => m.ActiveSelection()).Returns(() => activeSelection);
            selectionServiceMock.Setup(m => m.TrySetActiveSelection(It.IsAny<QualifiedSelection>()))
                .Returns(() => true).Callback((QualifiedSelection selection) => activeSelection = selection);
            return selectionServiceMock.Object;
        }

        public static T ParseAndTest<T>(IVBE vbe, Func<RubberduckParserState, T> testFunc)
        {
            T result = default;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                result = testFunc(state);
            }
            return result;
        }

        //public static T ParseAndTest<T>(IVBE vbe, Func<RubberduckParserState, IVBE, T> testFunc)
        //{
        //    T result = default;
        //    using (var state = MockParser.CreateAndParse(vbe))
        //    {
        //        result = testFunc(state, vbe);
        //    }
        //    return result;
        //}

        public static void ParseAndTest(IVBE vbe, Action<RubberduckParserState, IVBE, IRewritingManager> testFunc)
        {
            (RubberduckParserState state, IRewritingManager rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                testFunc(state, vbe, rewritingManager);
            }
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

        //public static T ParseAndTest<T>(Func<RubberduckParserState, IVBE, IRewritingManager, T> testFunc, TestMoveDefinition moveDefinition, string sourceContent)
        //{
        //    T result = default;
        //    var vbe = BuildVBEStub(moveDefinition, sourceContent);
        //    (RubberduckParserState state, IRewritingManager rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
        //    using (state)
        //    {
        //        result = testFunc(state, vbe, rewritingManager);
        //    }
        //    return result;
        //}

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

            var model = new MoveMemberModel(target, declarationFinderProvider, null, new MoveMemberObjectsFactory(declarationFinderProvider));

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
            var model = new MoveMemberModel(member, state, null, new MoveMemberObjectsFactory(state));

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

        public static IVBE BuildVBEStub(TestMoveDefinition moveDefinition, string sourceContent, string destinationContent = null, params ReferenceLibrary[] libraries)
        {
            if (moveDefinition.CreateNewModule)
            {
                moveDefinition.SetEndpointContent(sourceContent);
                return MockVbeBuilder.BuildFromModules(moveDefinition.ModuleDefinitions.Select(tc => tc.AsTuple)).Object;
            }
            moveDefinition.SetEndpointContent(sourceContent, destinationContent);
            return MockVbeBuilder.BuildFromModules(moveDefinition.ModuleDefinitions.Select(tc => tc.AsTuple), libraries).Object;
        }

        public static ModuleDefinition ClassModuleDefinition(string name, string content = null)
            => new ModuleDefinition(name, ComponentType.ClassModule, content);

        public static ModuleDefinition StdModuleDefinition(string name, string content = null)
            => new ModuleDefinition(name, ComponentType.StandardModule, content);

        public static ModuleDefinition UserFormDefinition(string name, string content = null)
            => new ModuleDefinition(name, ComponentType.UserForm, content);

        public static QualifiedModuleName GetQMN(IVBE VBE, string moduleName, string projectName = DEFAULT_PROJECT_NAME)
        {
            var project = VBE.VBProjects.Single(item => item.Name == projectName);
            var component = project.VBComponents.SingleOrDefault(c => c.Name == moduleName);
            using (component)
            {
                return component != null ? new QualifiedModuleName(component) : new QualifiedModuleName(project);
            }
        }

        public static string RetrieveModuleContent(IVBE vbe, string moduleName, string projectName = DEFAULT_PROJECT_NAME)
        {
            var vbProject = vbe.VBProjects.Single(item => item.Name == projectName);
            var component = vbProject.VBComponents.SingleOrDefault(item => item.Name == moduleName);
            using (component)
            {
                return component?.CodeModule.Content() ?? string.Empty;
            }
        }

        public static bool OccursOnce(string toFind, string content)
        {
            var firstIdx = content.IndexOf(toFind);
            var lastIdx = content.LastIndexOf(toFind);
            return firstIdx == lastIdx;
        }

        public static (string moduleName, ComponentType componentType) EndpointToSourceTuple(MoveEndpoints endpoints)
        {
            switch (endpoints)
            {
                case MoveEndpoints.FormToStd:
                    return (DEFAULT_SOURCE_FORM_NAME, ComponentType.UserForm);
                case MoveEndpoints.ClassToStd:
                    return (DEFAULT_SOURCE_CLASS_NAME, ComponentType.ClassModule);
                case MoveEndpoints.StdToStd:
                    return (DEFAULT_SOURCE_MODULE_NAME, ComponentType.StandardModule);
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        public static (string moduleName, ComponentType componentType) EndpointToDestinationTuple(MoveEndpoints endpoints)
        {
            switch (endpoints)
            {
                case MoveEndpoints.FormToStd:
                case MoveEndpoints.ClassToStd:
                case MoveEndpoints.StdToStd:
                    return (DEFAULT_DESTINATION_MODULE_NAME, ComponentType.StandardModule);
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private static Mock<IRefactoringPresenterFactory> CreatePresenterFactoryStub(MoveMemberModel model)
        {
            var presenterStub = new Mock<IMoveMemberPresenter>();
            presenterStub.Setup(p => p.Show()).Returns(model);

            var factoryStub = new Mock<IRefactoringPresenterFactory>();
            factoryStub.Setup(f => f.Create<IMoveMemberPresenter, MoveMemberModel>(It.IsAny<MoveMemberModel>()))
                .Returns(presenterStub.Object);

            return factoryStub;
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

    public struct ModuleDefinition
    {
        public ModuleDefinition(string name, ComponentType compType, string content = null)
        {
            ModuleName = name;
            ComponentType = compType;
            ModuleContent = content ?? $"{Tokens.Option} {Tokens.Explicit}";
        }

        public string ModuleName { get; }
        public ComponentType ComponentType { get; }
        public string ModuleContent { get; }
        public (string Name, string Content, ComponentType ComponentType) AsTuple
            => (ModuleName, ModuleContent, ComponentType);
    }

    public struct TestMoveDefinition
    {
        private List<ModuleDefinition> _moduleDefs;

        public MoveEndpoints Endpoints { get; }
        public string SelectedElement { get; }
        public DeclarationType SelectedDeclarationType { get; }

        private string _sourceModuleName;
        public string SourceModuleName => _sourceModuleName ?? DefaultSourceModuleNameForEndpoint(Endpoints);

        private string _destinationModuleName;
        public string DestinationModuleName => _destinationModuleName ?? DefaultDestinationModuleNameForEndpoint(Endpoints);

        public bool CreateNewModule { get; }

        public bool IsClassDestination => Endpoints == MoveEndpoints.ClassToClass || Endpoints == MoveEndpoints.StdToClass;
        public bool IsStdModuleDestination => Endpoints == MoveEndpoints.ClassToStd || Endpoints == MoveEndpoints.StdToStd;
        public bool IsClassSource => Endpoints == MoveEndpoints.ClassToClass || Endpoints == MoveEndpoints.ClassToStd;
        public bool IsFormSource => Endpoints == MoveEndpoints.FormToClass || Endpoints == MoveEndpoints.FormToStd;
        public bool IsStdModuleSource => Endpoints == MoveEndpoints.StdToClass || Endpoints == MoveEndpoints.StdToStd;

        public TestMoveDefinition(MoveEndpoints endpoints, string selectedElement = null, string sourceModuleName = null, string destinationModuleName = null, string sourceContent = null, string destinationContent = null, bool createNewModule = false)
            : this(endpoints, (selectedElement ?? string.Empty, DeclarationType.UnresolvedMember), sourceModuleName, destinationModuleName, sourceContent, destinationContent, createNewModule) { }

        public TestMoveDefinition(MoveEndpoints endpoints, (string, DeclarationType) selection, string sourceModuleName = null, string destinationModuleName = null, string sourceContent = null, string destinationContent = null, bool createNewModule = false)
        {
            _moduleDefs = new List<ModuleDefinition>();
            CreateNewModule = createNewModule;
            Endpoints = endpoints;
            SelectedElement = selection.Item1;
            SelectedDeclarationType = selection.Item2;

            _destinationModuleName = destinationModuleName;

            _sourceModuleName = sourceModuleName;

            if (sourceContent != null)
            {
                SetEndpointContent(sourceContent, destinationContent);
            }
        }

        public string DefaultSourceModuleNameForEndpoint(MoveEndpoints endpoints)
        {
            var defaultSourceModuleName = Support.DEFAULT_SOURCE_MODULE_NAME;
            switch (endpoints)
            {
                case MoveEndpoints.ClassToStd:
                    defaultSourceModuleName = Support.DEFAULT_SOURCE_CLASS_NAME;
                    break;
                case MoveEndpoints.FormToStd:
                    defaultSourceModuleName = Support.DEFAULT_SOURCE_FORM_NAME;
                    break;
            }
            return defaultSourceModuleName;
        }

        private string DefaultDestinationModuleNameForEndpoint(MoveEndpoints endpoints)
        {
            return IsStdModuleDestination 
                ? Support.DEFAULT_DESTINATION_MODULE_NAME
                : Support.DEFAULT_DESTINATION_CLASS_NAME;
        }

        public ComponentType DestinationComponentType
        {
            get
            {
                switch (Endpoints)
                {
                    case MoveEndpoints.ClassToClass:
                        return ComponentType.ClassModule;
                    case MoveEndpoints.StdToClass:
                        return ComponentType.ClassModule;
                    case MoveEndpoints.FormToClass:
                        return ComponentType.ClassModule;
                    default:
                        return ComponentType.StandardModule;
                }
            }
        }

        public ComponentType SourceComponentType
        {
            get
            {
                switch (Endpoints)
                {
                    case MoveEndpoints.ClassToStd:
                        return ComponentType.ClassModule;
                    case MoveEndpoints.ClassToClass:
                        return ComponentType.ClassModule;
                    case MoveEndpoints.FormToStd:
                        return ComponentType.UserForm;
                    case MoveEndpoints.FormToClass:
                        return ComponentType.UserForm;
                    default:
                        return ComponentType.StandardModule;
                }
            }
        }

        public void Add(ModuleDefinition moduleDef)
        {
            if (!_moduleDefs.Contains(moduleDef))
            {
                _moduleDefs.Add(moduleDef);
            }
        }

        public void SetEndpointContent(string sourceContent, string destinationContent = null)
        {
            Add(SourceModuleDefinition(sourceContent));
            if (!CreateNewModule)
            {
                Add(DestinationModuleDefinition(destinationContent));
            }
        }

        public ModuleDefinition[] ModuleDefinitions => _moduleDefs.ToArray();

        public ModuleDefinition SourceModuleDefinition(string content = null)
            => new ModuleDefinition(SourceModuleName, SourceComponentType, content ?? $"{Tokens.Option} {Tokens.Explicit}");

        public ModuleDefinition DestinationModuleDefinition(string content = null)
            => new ModuleDefinition(DestinationModuleName, DestinationComponentType, content ?? $"{Tokens.Option} {Tokens.Explicit}");

        public string ClassVariableName
            => $"{MoveMemberResources.Prefix_Variable}{DestinationModuleName}";

        public string ClassInstantiationSubName
        {
            get
            {
                if (SourceComponentType == ComponentType.ClassModule)
                {
                    return MoveMemberResources.Class_Initialize;
                }
                return $"{MoveMemberResources.Prefix_ClassInstantiationProcedure}{ClassVariableName}";
            }
        }
    }

    public struct MoveMemberTestResult
    {
        private Dictionary<string, string> _resultContent;
        private string _sourceModuleName;
        private string _destinationModuleName;

        public MoveMemberTestResult(TestMoveDefinition moveDefinition)
        {
            _sourceModuleName = moveDefinition.SourceModuleName;
            _destinationModuleName = moveDefinition.DestinationModuleName;
            _resultContent = new Dictionary<string, string>();
        }

        public void Add(string moduleName, string content)
        {
            _resultContent.Add(moduleName, content);
        }

        public string SourceContent
            => RetrieveContent(_sourceModuleName);

        public string DestinationContent
            => RetrieveContent(_destinationModuleName);

        public string RetrieveContent(string moduleName)
        {
            return _resultContent[moduleName];
        }
    }
}
