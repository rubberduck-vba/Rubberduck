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
        private readonly IMoveMemberMoveGroupsProviderFactory _moveGroupsProviderFactory;
        private readonly INameConflictFinder _nameConflictFinder;

        public MoveMemberStrategyFactory(
                RenameCodeDefinedIdentifierRefactoringAction renameAction,
                IMoveMemberMoveGroupsProviderFactory moveGroupsProviderFactory,
                INameConflictFinder nameConflictFinder)
        {
            _renameAction = renameAction;
            _moveGroupsProviderFactory = moveGroupsProviderFactory;
            _nameConflictFinder = nameConflictFinder;
        }

        public IMoveMemberRefactoringStrategy Create(MoveMemberStrategy strategyID)
        {
            switch (strategyID)
            {
                case MoveMemberStrategy.MoveToStandardModule:
                    return new MoveMemberToStdModule(_renameAction, _moveGroupsProviderFactory, _nameConflictFinder);
                default:
                    throw new ArgumentException();
            }
        }
    }
}
