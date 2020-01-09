using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
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
        IValidateVBAIdentifiers NameOnlyValidator(NameValidators validatorType);
        IEncapsulateFieldConflictFinder ConflictDetector(EncapsulateFieldStrategy strategy, IDeclarationFinderProvider declarationFinderProvider);
        void RegisterCandidates(IEnumerable<IEncapsulatableField> candidates);
    }

    public class EncapsulateFieldValidationsProvider : IEncapsulateFieldValidationsProvider
    {
        private Dictionary<NameValidators, IValidateVBAIdentifiers> _nameOnlyValidators;

        private List<IEncapsulatableField> _candidates;
        private List<IUserDefinedTypeMemberCandidate> _udtMemberCandidates;

        public EncapsulateFieldValidationsProvider()
        {
            _nameOnlyValidators = new Dictionary<NameValidators, IValidateVBAIdentifiers>()
            {
                [NameValidators.Default] = new IdentifierOnlyValidator(DeclarationType.Variable, false),
                [NameValidators.UserDefinedType] = new IdentifierOnlyValidator(DeclarationType.UserDefinedType, false),
                [NameValidators.UserDefinedTypeMember] = new IdentifierOnlyValidator(DeclarationType.UserDefinedTypeMember, false),
                [NameValidators.UserDefinedTypeMemberArray] = new IdentifierOnlyValidator(DeclarationType.UserDefinedTypeMember, true),
            };

            _candidates = new List<IEncapsulatableField>();
            _udtMemberCandidates = new List<IUserDefinedTypeMemberCandidate>();
        }

        public IValidateVBAIdentifiers NameOnlyValidator(NameValidators validatorType)
            => _nameOnlyValidators[validatorType];

        public IEncapsulateFieldConflictFinder ConflictDetector(EncapsulateFieldStrategy strategy, IDeclarationFinderProvider declarationFinderProvider)
        {
            if (strategy == EncapsulateFieldStrategy.UseBackingFields)
            {
                return new UseBackingFieldsConflictFinder(declarationFinderProvider, _candidates, _udtMemberCandidates);
            }
            return new ConvertFieldsToUDTMembersConflictFinder(declarationFinderProvider, _candidates, _udtMemberCandidates);
        }

        public void RegisterCandidates(IEnumerable<IEncapsulatableField> candidates)
        {
            _candidates.AddRange(candidates);
            foreach (var udtCandidate in candidates.Where(c => c is IUserDefinedTypeCandidate).Cast<IUserDefinedTypeCandidate>())
            {
                foreach (var member in udtCandidate.Members)
                {
                    _udtMemberCandidates.Add(member);
                }
            }
        }
    }
}
