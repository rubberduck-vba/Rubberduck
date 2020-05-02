using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.Refactorings.Rename;
using System;
using System.Collections.Generic;

namespace Rubberduck.Refactorings
{
    public interface IMoveMemberStrategyFactory
    {
        IEnumerable<IMoveMemberRefactoringStrategy> CreateAll();
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

        public IEnumerable<IMoveMemberRefactoringStrategy> CreateAll()
        {
            var allStrategies = new List<IMoveMemberRefactoringStrategy>();
            var anySourceToStdModule = new MoveMemberToStdModule(_declarationFinderProvider, _renameAction, _moveGroupsProviderFactory, _nameConflictResolverFactory, _declarationProxyFactory);

            allStrategies.Add(anySourceToStdModule);
            return allStrategies;
        }
    }
}
