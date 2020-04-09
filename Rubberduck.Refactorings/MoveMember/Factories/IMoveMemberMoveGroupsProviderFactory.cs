using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveMember;
using System.Collections.Generic;

namespace Rubberduck.Refactorings
{
    public interface IMoveMemberMoveGroupsProviderFactory
    {
        IMoveMemberGroupsProvider Create(IEnumerable<IMoveableMemberSet> moveableMemberSets);
    }

    public class MoveMemberMoveGroupsProviderFactory : IMoveMemberMoveGroupsProviderFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        public MoveMemberMoveGroupsProviderFactory(IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public IMoveMemberGroupsProvider Create(IEnumerable<IMoveableMemberSet> moveableMemberSets)
        {
            return new MoveMemberGroupsProvider(moveableMemberSets, _declarationFinderProvider);
        }
    }
}
