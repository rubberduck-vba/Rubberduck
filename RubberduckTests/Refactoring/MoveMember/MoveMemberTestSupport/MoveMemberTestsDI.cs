using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.Refactorings.Rename;
using Rubberduck.VBEditor.Utility;
using System;

namespace RubberduckTests.Refactoring.MoveMember
{

    public class MoveMemberTestsDI
    {
        public static T Resolve<T>(RubberduckParserState state, IRewritingManager rewritingManager) where T: class
        {
            return Resolve<T>(state, rewritingManager, typeof(T).Name);
        }

        private static T Resolve<T>(RubberduckParserState _state, IRewritingManager _rewritingManager, string name) where T : class
        {
            switch (name)
            {
                case nameof(MoveMemberRefactoringAction):
                    return new MoveMemberRefactoringAction(
                            Resolve<MoveMemberToNewStandardModuleRefactoringAction>(_state, _rewritingManager), 
                            Resolve<MoveMemberToExistingStandardModuleRefactoringAction>(_state, _rewritingManager)) as T;
                case nameof(MoveMemberToNewStandardModuleRefactoringAction):
                    return new MoveMemberToNewStandardModuleRefactoringAction(
                            _state, 
                            _rewritingManager,
                            Resolve<MovedContentProviderFactory>(_state, _rewritingManager),
                            Resolve<IMoveMemberStrategyFactory>(_state, _rewritingManager),
                            Resolve<IAddComponentService>(_state, _rewritingManager)) as T;
                case nameof(IAddComponentService):
                    return MoveMemberRefactoringActionTestSupportBase.TestAddComponentService(_state?.ProjectsProvider) as T;
                case nameof(IMoveMemberRefactoringPreviewerFactory):
                    return new MoveMemberRefactoringPreviewerFactory(
                            Resolve<MoveMemberToExistingStandardModuleRefactoringAction>(_state, _rewritingManager),
                            _rewritingManager,
                            Resolve<MovedContentProviderFactory>(_state, _rewritingManager)) as T;
                case nameof(MoveMemberRefactoringPreviewerFactory):
                    return new MoveMemberRefactoringPreviewerFactory(
                            Resolve<MoveMemberToExistingStandardModuleRefactoringAction>(_state, _rewritingManager),
                            _rewritingManager,
                            Resolve<MovedContentProviderFactory>(_state, _rewritingManager)) as T;
                case nameof(MoveMemberToExistingStandardModuleRefactoringAction):
                    return new MoveMemberToExistingStandardModuleRefactoringAction(
                            _rewritingManager, 
                            Resolve<MovedContentProviderFactory>(_state, _rewritingManager),
                            Resolve<IMoveMemberStrategyFactory>(_state, _rewritingManager)
                            ) as T;
                case nameof(MoveMemberStrategyFactory):
                    return new MoveMemberStrategyFactory(
                            Resolve<RenameCodeDefinedIdentifierRefactoringAction>(_state, _rewritingManager),
                            Resolve<MoveMemberMoveGroupsProviderFactory>(_state, _rewritingManager),
                            Resolve<INameConflictFinder>(_state, _rewritingManager)) as T;
                case nameof(IMoveMemberStrategyFactory):
                    return new MoveMemberStrategyFactory(
                            Resolve<RenameCodeDefinedIdentifierRefactoringAction>(_state, _rewritingManager),
                            Resolve<MoveMemberMoveGroupsProviderFactory>(_state, _rewritingManager),
                            Resolve<INameConflictFinder>(_state, _rewritingManager)) as T;
                case nameof(RenameCodeDefinedIdentifierRefactoringAction):
                    return new RenameCodeDefinedIdentifierRefactoringAction(
                            _state, 
                            _state?.ProjectsProvider, 
                            _rewritingManager) as T;
                case nameof(MoveMemberEndpointFactory):
                    return new MoveMemberEndpointFactory(
                            _state, Resolve<MoveableMemberSetsFactory>(_state, _rewritingManager)) as T;
                case nameof(MoveMemberMoveGroupsProviderFactory):
                    return new MoveMemberMoveGroupsProviderFactory(_state) as T;
                case nameof(MovedContentPreviewProvider):
                    return new MovedContentPreviewProvider() as T;
                case nameof(MovedContentProvider):
                    return new MovedContentProvider() as T;
                case nameof(MovedContentProviderFactory):
                    return new MovedContentProviderFactory() as T;
                case nameof(IMoveableMemberSetsFactory):
                    return new MoveableMemberSetsFactory(
                            _state, Resolve<MoveableMemberSetFactory>(_state, _rewritingManager)) as T;
                case nameof(MoveableMemberSetsFactory):
                    return new MoveableMemberSetsFactory(
                            _state, Resolve<MoveableMemberSetFactory>(_state, _rewritingManager)) as T;
                case nameof(MoveableMemberSetFactory):
                    return new MoveableMemberSetFactory() as T;
                case nameof(INameConflictFinder):
                    return new NameConflictFinder(_state) as T;
                default:
                    throw new ArgumentException();
            }
        }
    }
}
