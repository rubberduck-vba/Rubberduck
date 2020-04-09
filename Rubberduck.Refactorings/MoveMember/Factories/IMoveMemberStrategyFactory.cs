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

        public MoveMemberStrategyFactory(
                RenameCodeDefinedIdentifierRefactoringAction renameAction,
                IMoveMemberMoveGroupsProviderFactory moveGroupsProviderFactory)
        {
            _renameAction = renameAction;
            _moveGroupsProviderFactory = moveGroupsProviderFactory;
        }

        public IMoveMemberRefactoringStrategy Create(MoveMemberStrategy strategyID)
        {
            switch (strategyID)
            {
                case MoveMemberStrategy.MoveToStandardModule:
                    return new MoveMemberToStdModule(_renameAction, _moveGroupsProviderFactory);
                default:
                    throw new ArgumentException();
            }
        }
    }
}
