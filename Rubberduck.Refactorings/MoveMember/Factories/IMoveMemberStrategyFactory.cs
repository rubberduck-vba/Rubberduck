using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.Refactorings.Rename;
using System;

namespace Rubberduck.Refactorings
{
    public enum MoveMemberStrategy
    {
        MoveToStandardModule,
    }

    public interface IMoveMemberStrategyFactory
    {
        IMoveMemberRefactoringStrategy Create(MoveMemberStrategy strategyID);
    }

    public class MoveMemberStrategyFactory : IMoveMemberStrategyFactory
    {
        private readonly RenameCodeDefinedIdentifierRefactoringAction _renameAction;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IMoveMemberMoveGroupsProviderFactory _moveGroupsProviderFactory;
        private readonly INameConflictFinder _nameConflictFinder;
        private readonly IDeclarationProxyFactory _declarationProxyFactory;

        public MoveMemberStrategyFactory(
                IDeclarationFinderProvider declarationFinderProvider,
                RenameCodeDefinedIdentifierRefactoringAction renameAction,
                IMoveMemberMoveGroupsProviderFactory moveGroupsProviderFactory,
                INameConflictFinder nameConflictFinder,
                IDeclarationProxyFactory declarationProxyFactory)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _renameAction = renameAction;
            _moveGroupsProviderFactory = moveGroupsProviderFactory;
            _nameConflictFinder = nameConflictFinder;
            _declarationProxyFactory = declarationProxyFactory;
        }

        public IMoveMemberRefactoringStrategy Create(MoveMemberStrategy strategyID)
        {
            switch (strategyID)
            {
                case MoveMemberStrategy.MoveToStandardModule:
                    return new MoveMemberToStdModule(_declarationFinderProvider, _renameAction, _moveGroupsProviderFactory, _nameConflictFinder, _declarationProxyFactory);
                default:
                    throw new ArgumentException();
            }
        }
    }
}
