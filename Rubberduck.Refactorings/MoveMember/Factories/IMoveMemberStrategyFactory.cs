using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.Refactorings.Rename;
using System;

namespace Rubberduck.Refactorings
{
    public interface IMoveMemberStrategyFactory
    {
        IMoveMemberRefactoringStrategy Create(MoveEndpoints moveEndpoints);
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

        public IMoveMemberRefactoringStrategy Create(MoveEndpoints moveEndpoints)
        {
            switch (moveEndpoints)
            {
                case MoveEndpoints.StdToStd:
                case MoveEndpoints.ClassToStd:
                case MoveEndpoints.FormToStd:
                    return new MoveMemberToStdModule(_declarationFinderProvider, _renameAction, _moveGroupsProviderFactory, _nameConflictFinder, _declarationProxyFactory);
                case MoveEndpoints.StdToClass:
                case MoveEndpoints.FormToClass:
                case MoveEndpoints.ClassToClass:
                    return new MoveMemberStdToClassModule(_declarationFinderProvider, _renameAction, _moveGroupsProviderFactory, _nameConflictFinder, _declarationProxyFactory);
                default:
                    throw new ArgumentException();
            }
        }
    }
}
