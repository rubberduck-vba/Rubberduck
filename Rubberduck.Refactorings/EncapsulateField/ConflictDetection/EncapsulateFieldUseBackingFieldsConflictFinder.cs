using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldUseBackingFieldsConflictFinder : EncapsulateFieldConflictFinderBase, IEncapsulateFieldConflictFinder
    {
        public EncapsulateFieldUseBackingFieldsConflictFinder(IDeclarationFinderProvider declarationFinderProvider, IEnumerable<IEncapsulateFieldCandidate> candidates)
            : base(declarationFinderProvider, candidates)
        { }

        protected override IEnumerable<Declaration> FindRelevantMembers(IEncapsulateFieldCandidate candidate)
            => _declarationFinderProvider.DeclarationFinder.Members(candidate.QualifiedModuleName)
                .Where(d => d != candidate.Declaration);
    }
}
