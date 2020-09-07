using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldUseBackingUDTMemberConflictFinder : IEncapsulateFieldConflictFinder
    {
        IObjectStateUDT AssignNoConflictIdentifiers(IObjectStateUDT stateUDT, IDeclarationFinderProvider declarationFinderProvider);
    }

    public class EncapsulateFieldUseBackingUDTMemberConflictFinder : EncapsulateFieldConflictFinderBase, IEncapsulateFieldUseBackingUDTMemberConflictFinder
    {
        private static DeclarationType[] _udtTypeIdentifierNonConflictTypes = new DeclarationType[]
        {
            DeclarationType.Project,
            DeclarationType.Module,
            DeclarationType.Property,
            DeclarationType.Function,
            DeclarationType.Procedure,
            DeclarationType.Variable,
            DeclarationType.Constant,
            DeclarationType.UserDefinedTypeMember,
            DeclarationType.EnumerationMember,
            DeclarationType.Parameter
        };

        private List<IObjectStateUDT> _objectStateUDTs;
        public EncapsulateFieldUseBackingUDTMemberConflictFinder(IDeclarationFinderProvider declarationFinderProvider, IEnumerable<IEncapsulateFieldCandidate> candidates, IEnumerable<IObjectStateUDT> objectStateUDTs)
            : base(declarationFinderProvider, candidates)
        {
            _objectStateUDTs = objectStateUDTs.ToList();
        }

        public override bool TryValidateEncapsulationAttributes(IEncapsulateFieldCandidate field, out string errorMessage)
        {
            errorMessage = string.Empty;
            if (!field.EncapsulateFlag)
            {
                return true;
            }

            if (!base.TryValidateEncapsulationAttributes(field, out errorMessage))
            {
                return false;
            }

            //Compare to existing members...they cannot change
            var objectStateUDT = _objectStateUDTs.SingleOrDefault(os => os.IsSelected);
            return !ConflictsWithExistingUDTMembers(objectStateUDT, field.BackingIdentifier);
        }

        public IObjectStateUDT AssignNoConflictIdentifiers(IObjectStateUDT stateUDT, IDeclarationFinderProvider declarationFinderProvider)
        {
            var members = declarationFinderProvider.DeclarationFinder.Members(stateUDT.QualifiedModuleName);
            var guard = 0;
            while (guard++ < 10 && members.Any(m => m.IdentifierName.IsEquivalentVBAIdentifierTo(stateUDT.FieldIdentifier)))
            {
                stateUDT.FieldIdentifier = stateUDT.FieldIdentifier.IncrementEncapsulationIdentifier();
            }

            members = declarationFinderProvider.DeclarationFinder.Members(stateUDT.QualifiedModuleName)
                .Where(m => !_udtTypeIdentifierNonConflictTypes.Any(nct => m.DeclarationType.HasFlag(nct)));

            guard = 0;
            while (guard++ < 10 && members.Any(m => m.IdentifierName.IsEquivalentVBAIdentifierTo(stateUDT.TypeIdentifier)))
            {
                stateUDT.TypeIdentifier = stateUDT.TypeIdentifier.IncrementEncapsulationIdentifier();
            }
            return stateUDT;
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
            if (objectStateUDT is null)
            {
                return false;
            }

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
