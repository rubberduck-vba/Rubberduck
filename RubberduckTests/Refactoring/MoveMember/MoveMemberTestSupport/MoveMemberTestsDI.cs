using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.Refactorings.Rename;
using Rubberduck.VBEditor.Utility;
using System;

namespace RubberduckTests.Refactoring.MoveMember
{

    public class MoveMemberTestsDI
    {
        private static RubberduckParserState _state;
        private static IRewritingManager _rewritingManager;

        public static void Initialize(RubberduckParserState state, IRewritingManager rewritingManager)
        {
            _state = state;
            _rewritingManager = rewritingManager;
        }

        public static T Resolve<T>() where T : class
        {
            return Resolve<T>(typeof(T).Name);
        }

        public static T Resolve<T>(string name) where T : class
        {
            switch (name)
            {
                case nameof(MoveMemberRefactoringAction):
                    return new MoveMemberRefactoringAction(
                            Resolve<MoveMemberToNewModuleRefactoringAction>(), 
                            Resolve<MoveMemberExistingModulesRefactoringAction>()) as T;
                case nameof(MoveMemberToNewModuleRefactoringAction):
                    return new MoveMemberToNewModuleRefactoringAction(
                            Resolve<MoveMemberExistingModulesRefactoringAction>(), 
                            _state, 
                            _rewritingManager, 
                            Resolve<IAddComponentService>()) as T;
                case nameof(IAddComponentService):
                    return MoveMemberRefactoringActionTestSupportBase.TestAddComponentService(_state?.ProjectsProvider) as T;
                case nameof(IMoveMemberRefactoringPreviewerFactory):
                    return new MoveMemberRefactoringPreviewerFactory(
                            Resolve<MoveMemberExistingModulesRefactoringAction>(),
                            _rewritingManager,
                            Resolve<MovedContentProviderFactory>()) as T;
                case nameof(MoveMemberRefactoringPreviewerFactory):
                    return new MoveMemberRefactoringPreviewerFactory(
                            Resolve<MoveMemberExistingModulesRefactoringAction>(),
                            _rewritingManager,
                            Resolve<MovedContentProviderFactory>()) as T;
                case nameof(MoveMemberExistingModulesRefactoringAction):
                    return new MoveMemberExistingModulesRefactoringAction(
                            _rewritingManager, 
                            Resolve<MovedContentProviderFactory>()) as T;
                case nameof(MoveMemberStrategyFactory):
                    return new MoveMemberStrategyFactory(
                            Resolve<RenameCodeDefinedIdentifierRefactoringAction>(), 
                            Resolve<MoveMemberMoveGroupsProviderFactory>()) as T;
                case nameof(RenameCodeDefinedIdentifierRefactoringAction):
                    return new RenameCodeDefinedIdentifierRefactoringAction(
                            _state, 
                            _state?.ProjectsProvider, 
                            _rewritingManager) as T;
                case nameof(MoveMemberEndpointFactory):
                    return new MoveMemberEndpointFactory(
                            _state, Resolve<MoveableMemberSetsFactory>()) as T;
                case nameof(MoveMemberMoveGroupsProviderFactory):
                    return new MoveMemberMoveGroupsProviderFactory(_state) as T;
                case nameof(MovedContentPreviewProvider):
                    return new MovedContentPreviewProvider() as T;
                case nameof(MovedContentProvider):
                    return new MovedContentProvider() as T;
                case nameof(MovedContentProviderFactory):
                    return new MovedContentProviderFactory() as T;
                case nameof(MoveableMemberSetsFactory):
                    return new MoveableMemberSetsFactory(
                            _state, Resolve<MoveableMemberSetFactory>()) as T;
                case nameof(MoveableMemberSetFactory):
                    return new MoveableMemberSetFactory() as T;
                default:
                    throw new ArgumentException();
            }
        }
    }
}
