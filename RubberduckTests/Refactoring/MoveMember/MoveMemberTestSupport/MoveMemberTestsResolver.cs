using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.Refactorings.Rename;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SourceCodeHandling;
using Rubberduck.VBEditor.Utility;
using System;
using System.Linq;

namespace RubberduckTests.Refactoring.MoveMember
{

    public class MoveMemberTestsResolver
    {
        private readonly RubberduckParserState _state;
        private readonly IRewritingManager _rewritingManager;
        public MoveMemberTestsResolver(RubberduckParserState state, IRewritingManager rewritingManager = null)
        {
            _state = state;
            _rewritingManager = rewritingManager;
        }

        public static MoveMemberRefactoring CreateRefactoring(IRewritingManager rewritingManager, RubberduckParserState state, RefactoringUserInteraction<IMoveMemberPresenter, MoveMemberModel> userInteraction, ISelectionService selectionService)
        {
            var serviceLocator = new MoveMemberTestsResolver(state, rewritingManager);
            var selectedDeclarationService = new SelectedDeclarationProvider(selectionService, state);
            return new MoveMemberRefactoring(serviceLocator.Resolve<MoveMemberRefactoringAction>(),
                                                userInteraction,
                                                selectionService,
                                                selectedDeclarationService,
                                                serviceLocator.Resolve<IMoveMemberModelFactory>()
                                                );
        }

        public static MoveMemberModel CreateRefactoringModel(string identifier, DeclarationType declarationType, RubberduckParserState state)
        {
            var serviceLocator = new MoveMemberTestsResolver(state);
            var target = state.DeclarationFinder.DeclarationsWithType(declarationType).Where(d => d.IdentifierName == identifier).Single();
            return serviceLocator.Resolve<IMoveMemberModelFactory>().Create(target);
        }

        public static MoveMemberModel CreateRefactoringModel(Declaration target, IDeclarationFinderProvider declarationFinderProvider)
        {
            var serviceLocator = new MoveMemberTestsResolver(declarationFinderProvider as RubberduckParserState);
            return serviceLocator.Resolve<IMoveMemberModelFactory>().Create(target);
        }

        public T Resolve<T>() where T : class
        {
            return Resolve<T>(typeof(T).Name);
        }

        private T Resolve<T>(string name) where T : class
        {
            switch (name)
            {
                case nameof(MoveMemberRefactoringAction):
                    return new MoveMemberRefactoringAction(
                            Resolve<MoveMemberToNewModuleRefactoringAction>(),
                            Resolve<MoveMemberToExistingModuleRefactoringAction>()) as T;
                case nameof(MoveMemberToNewModuleRefactoringAction):
                    return new MoveMemberToNewModuleRefactoringAction(
                            _state,
                            _rewritingManager,
                            Resolve<IMovedContentProviderFactory>(),
                            Resolve<IMoveMemberStrategyFactory>(),
                            Resolve<IAddComponentService>()) as T;
                case nameof(IAddComponentService):
                    return TestAddComponentService(_state?.ProjectsProvider) as T;
                case nameof(IMoveMemberRefactoringPreviewerFactory):
                    return new MoveMemberRefactoringPreviewerFactory(
                            Resolve<MoveMemberToExistingModuleRefactoringAction>(),
                            _rewritingManager,
                            Resolve<IMovedContentProviderFactory>()) as T;
                case nameof(MoveMemberToExistingModuleRefactoringAction):
                    return new MoveMemberToExistingModuleRefactoringAction(
                            _rewritingManager,
                            Resolve<IMovedContentProviderFactory>(),
                            Resolve<IMoveMemberStrategyFactory>()
                            ) as T;
                case nameof(IMoveMemberStrategyFactory):
                    return new MoveMemberStrategyFactory(
                            Resolve<IDeclarationFinderProvider>(),
                            Resolve<RenameCodeDefinedIdentifierRefactoringAction>(),
                            Resolve<IMoveMemberMoveGroupsProviderFactory>(),
                            Resolve<INameConflictFinder>(),
                            Resolve<IDeclarationProxyFactory>()) as T;
                case nameof(RenameCodeDefinedIdentifierRefactoringAction):
                    return new RenameCodeDefinedIdentifierRefactoringAction(
                            _state,
                            _state?.ProjectsProvider,
                            _rewritingManager) as T;
                case nameof(IMoveMemberEndpointFactory):
                    return new MoveMemberEndpointFactory(
                            _state, Resolve<IMoveableMemberSetsFactory>()) as T;
                case nameof(IMoveMemberMoveGroupsProviderFactory):
                    return new MoveMemberMoveGroupsProviderFactory(_state) as T;
                case nameof(IMovedContentPreviewProvider):
                    return new MovedContentPreviewProvider() as T;
                case nameof(IMovedContentProviderFactory):
                    return new MovedContentProviderFactory() as T;
                case nameof(IMoveableMemberSetsFactory):
                    return new MoveableMemberSetsFactory(
                            _state, Resolve<MoveableMemberSetFactory>()) as T;
                case nameof(MoveableMemberSetFactory):
                    return new MoveableMemberSetFactory() as T;
                case nameof(INameConflictFinder):
                    return new NameConflictFinder(_state, Resolve<IDeclarationProxyFactory>()) as T;
                case nameof(IDeclarationProxyFactory):
                    return new DeclarationProxyFactory(_state) as T;
                case nameof(IDeclarationFinderProvider):
                    return _state as T;
                case nameof(IMoveMemberModelFactory):
                    return new MoveMemberModelFactory(_state, Resolve<IMoveMemberStrategyFactory>(), Resolve<IMoveMemberEndpointFactory>()) as T;
                default:
                    throw new ArgumentException();
            }
        }

        private static IAddComponentService TestAddComponentService(IProjectsProvider projectsProvider)
        {
            var sourceCodeHandler = new CodeModuleComponentSourceCodeHandler();
            return new AddComponentService(projectsProvider, sourceCodeHandler, sourceCodeHandler);
        }
    }
}
