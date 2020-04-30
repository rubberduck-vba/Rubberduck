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
        private readonly IConflictDetectionSessionFactory _nameConflictResolverFactory;
        private readonly IConflictDetectionDeclarationProxyFactory _declarationProxyFactory;

        public MoveMemberStrategyFactory(
                IDeclarationFinderProvider declarationFinderProvider,
                RenameCodeDefinedIdentifierRefactoringAction renameAction,
                IMoveMemberMoveGroupsProviderFactory moveGroupsProviderFactory,
                IConflictDetectionSessionFactory nameConflictResolverFactory,
                IConflictDetectionDeclarationProxyFactory declarationProxyFactory)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _renameAction = renameAction;
            _moveGroupsProviderFactory = moveGroupsProviderFactory;
            _nameConflictResolverFactory = nameConflictResolverFactory;
            _declarationProxyFactory = declarationProxyFactory;
        }

        public IMoveMemberRefactoringStrategy Create(MoveEndpoints moveEndpoints)
        {
            switch (moveEndpoints)
            {
                case MoveEndpoints.StdToStd:
                case MoveEndpoints.ClassToStd:
                case MoveEndpoints.FormToStd:
                    return new MoveMemberToStdModule(_declarationFinderProvider, _renameAction, _moveGroupsProviderFactory, _nameConflictResolverFactory, _declarationProxyFactory);
                case MoveEndpoints.StdToClass:
                case MoveEndpoints.FormToClass:
                case MoveEndpoints.ClassToClass:
                    return new MoveMemberStdToClassModule(_declarationFinderProvider, _renameAction, _moveGroupsProviderFactory, _nameConflictResolverFactory, _declarationProxyFactory);
                default:
                    throw new ArgumentException();
            }
        }
    }
}
