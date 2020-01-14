using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class ConvertFieldsToUDTMembersStrategyConflictFinder : EncapsulateFieldConflictFinderBase
    {
        private IEnumerable<IObjectStateUDT> _objectStateUDTs;
        public ConvertFieldsToUDTMembersStrategyConflictFinder(IDeclarationFinderProvider declarationFinderProvider, IEnumerable<IEncapsulateFieldCandidate> candidates, IEnumerable<IUserDefinedTypeMemberCandidate> udtCandidates, IEnumerable<IObjectStateUDT> objectStateUDTs)
            : base(declarationFinderProvider, candidates, udtCandidates)
        {
            _objectStateUDTs = objectStateUDTs;
        }

        public override bool TryValidateEncapsulationAttributes(IEncapsulateFieldCandidate field, out string errorMessage)
        {
            errorMessage = string.Empty;
            if (!field.EncapsulateFlag) { return true; }

            if (!base.TryValidateEncapsulationAttributes(field, out errorMessage))
            {
                return false;
            }

            //Compare to existing members...they cannot change
            var objectStateUDT = _objectStateUDTs.SingleOrDefault(os => os.IsSelected);
            return !ConflictsWithExistingUDTMembers(objectStateUDT, field.BackingIdentifier);
        }

        public override IEncapsulateFieldCandidate AssignNoConflictIdentifiers(IEncapsulateFieldCandidate candidate)
        {
            candidate = base.AssignNoConflictIdentifier(candidate, DeclarationType.Property);

            var objectStateUDT = _objectStateUDTs.SingleOrDefault(os => os.IsSelected);
            var guard = 0;
            while (guard++ < 10 && ConflictsWithExistingUDTMembers(objectStateUDT, candidate.PropertyIdentifier))
            {
                candidate.PropertyIdentifier = candidate.PropertyIdentifier.IncrementEncapsulationIdentifier();
            }
            return candidate;
        }

        protected override IEncapsulateFieldCandidate AssignNoConflictIdentifier(IEncapsulateFieldCandidate candidate, DeclarationType declarationType)
        {
            candidate = base.AssignNoConflictIdentifier(candidate, declarationType);

            var objectStateUDT = _objectStateUDTs.SingleOrDefault(os => os.IsSelected);
            var guard = 0;
            while (guard++ < 10 && ConflictsWithExistingUDTMembers(objectStateUDT, candidate.BackingIdentifier))
            {
                candidate.BackingIdentifier = candidate.BackingIdentifier.IncrementEncapsulationIdentifier();
            }
            return candidate;
        }

        private bool ConflictsWithExistingUDTMembers(IObjectStateUDT objectStateUDT, string identifier)
        {
            if (objectStateUDT is null) { return false; }

            return objectStateUDT.ExistingMembers.Any(nm => nm.IdentifierName.IsEquivalentVBAIdentifierTo(identifier));
        }

        protected override IEnumerable<Declaration> FindRelevantMembers(IEncapsulateFieldCandidate candidate)
        {
            var members = _declarationFinderProvider.DeclarationFinder.Members(candidate.QualifiedModuleName)
                .Where(d => d != candidate.Declaration);

            var membersToRemove = _fieldCandidates.Where(fc => fc.EncapsulateFlag && fc.Declaration.DeclarationType.HasFlag(DeclarationType.Variable))
                .Select(fc => fc.Declaration);

            return members.Except(membersToRemove);
        }
    }
}
