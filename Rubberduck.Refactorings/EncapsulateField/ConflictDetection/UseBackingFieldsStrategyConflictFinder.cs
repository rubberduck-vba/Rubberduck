using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class UseBackingFieldsStrategyConflictFinder : EncapsulateFieldConflictFinderBase, IEncapsulateFieldConflictFinder
    {
        public UseBackingFieldsStrategyConflictFinder(IDeclarationFinderProvider declarationFinderProvider, IEnumerable<IEncapsulateFieldCandidate> candidates, IEnumerable<IUserDefinedTypeMemberCandidate> udtCandidates)
            : base(declarationFinderProvider, candidates, udtCandidates) { }

        protected override IEnumerable<Declaration> FindRelevantMembers(IEncapsulateFieldCandidate candidate)
            => _declarationFinderProvider.DeclarationFinder.Members(candidate.QualifiedModuleName)
                .Where(d => d != candidate.Declaration);
    }
}
