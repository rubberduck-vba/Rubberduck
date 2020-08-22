using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
using Rubberduck.Resources;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public enum NameValidators
    {
        Default,
        UserDefinedType,
        UserDefinedTypeMember,
        UserDefinedTypeMemberArray
    }

    public interface IEncapsulateFieldValidationsProvider
    {
        IEncapsulateFieldConflictFinder ConflictDetector(EncapsulateFieldStrategy strategy, IDeclarationFinderProvider declarationFinderProvider);
    }

    public class EncapsulateFieldValidationsProvider : IEncapsulateFieldValidationsProvider
    {
        private static Dictionary<NameValidators, IValidateVBAIdentifiers> _nameOnlyValidators = new Dictionary<NameValidators, IValidateVBAIdentifiers>()
        {
            [NameValidators.Default] = new IdentifierOnlyValidator(DeclarationType.Variable, false),
            [NameValidators.UserDefinedType] = new IdentifierOnlyValidator(DeclarationType.UserDefinedType, false),
            [NameValidators.UserDefinedTypeMember] = new IdentifierOnlyValidator(DeclarationType.UserDefinedTypeMember, false),
            [NameValidators.UserDefinedTypeMemberArray] = new IdentifierOnlyValidator(DeclarationType.UserDefinedTypeMember, true),
        };

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


        private List<IEncapsulateFieldCandidate> _candidates;
        private List<IUserDefinedTypeMemberCandidate> _udtMemberCandidates;
        private List<IObjectStateUDT> _objectStateUDTs;

        public EncapsulateFieldValidationsProvider(IEnumerable<IEncapsulateFieldCandidate> candidates, IEnumerable<IObjectStateUDT> objectStateUDTCandidates)
        {
            _udtMemberCandidates = new List<IUserDefinedTypeMemberCandidate>();
            _objectStateUDTs = objectStateUDTCandidates.ToList();
            _candidates = candidates.ToList();
            var udtCandidates = candidates.Where(c => c is IUserDefinedTypeCandidate).Cast<IUserDefinedTypeCandidate>();

            foreach (var udtCandidate in candidates.Where(c => c is IUserDefinedTypeCandidate).Cast<IUserDefinedTypeCandidate>())
            {
                LoadUDTMemberCandidates(udtCandidate);
            }
        }

        private void LoadUDTMemberCandidates(IUserDefinedTypeCandidate udtCandidate)
        {
            foreach (var member in udtCandidate.Members)
            {
                if (member.WrappedCandidate is IUserDefinedTypeCandidate udt)
                {
                    LoadUDTMemberCandidates(udt);
                }
                _udtMemberCandidates.Add(member);
            }
        }

        public static IValidateVBAIdentifiers NameOnlyValidator(NameValidators validatorType)
            => _nameOnlyValidators[validatorType];

        public static IObjectStateUDT AssignNoConflictIdentifiers(IObjectStateUDT stateUDT, IDeclarationFinderProvider declarationFinderProvider)
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

        public IEncapsulateFieldConflictFinder ConflictDetector(EncapsulateFieldStrategy strategy, IDeclarationFinderProvider declarationFinderProvider)
        {
            if (strategy == EncapsulateFieldStrategy.UseBackingFields)
            {
                return new UseBackingFieldsStrategyConflictFinder(declarationFinderProvider,  _candidates,  _udtMemberCandidates);
            }
            return new ConvertFieldsToUDTMembersStrategyConflictFinder(declarationFinderProvider, _candidates, _udtMemberCandidates, _objectStateUDTs);
        }
    }
}
