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
    public interface IValidateVBAIdentifiers
    {
        bool IsValidVBAIdentifier(string identifier, out string errorMessage);
        bool IsValidUDTMemberIdentifier(string identifier, bool isArray, out string errorMessage);
    }

    public enum Validators
    {
        Default,
        UserDefinedType,
        UserDefinedTypeMember,
        UserDefinedTypeMemberArray
    }

    public interface IValidateEncapsulateFieldNames
    {
        bool IsValidVBAIdentifier(string identifier, DeclarationType declarationType, out string errorMessage, bool isArray = false);
        bool HasConflictingIdentifier(IEncapsulateFieldCandidate candidate, DeclarationType declarationType, out string errorMessage);
        bool HasConflictingIdentifierIgnoreEncapsulationFlag(IEncapsulateFieldCandidate field, DeclarationType declarationType, out string errorMessage);
        bool IsConflictingProposedIdentifier(string fieldName, IEncapsulateFieldCandidate candidate, DeclarationType declarationType);
        IEncapsulateFieldCandidate AssignNoConflictIdentifier(IEncapsulateFieldCandidate candidate, DeclarationType declarationType);
    }

    //public interface IEncapsulateFieldValidator : IValidateEncapsulateFieldNames
    //{
    //    string CreateNonConflictIdentifierForProposedType(string identifier, QualifiedModuleName qmn, DeclarationType declarationType);
    //}

    public interface IEncapsulateFieldValidationsProvider
    {
        IValidateVBAIdentifiers NameOnlyValidator(Validators validatorType);
        IEncapsulateFieldConflictFinder ConflictDetector(EncapsulateFieldStrategy strategy, IDeclarationFinderProvider declarationFinderProvider);
        void RegisterCandidates(IEnumerable<IEncapsulateFieldCandidate> candidates);
        void RegisterCandidate(IEncapsulateFieldCandidate candidate);
    }

    public class EncapsulateFieldValidationsProvider : IEncapsulateFieldValidationsProvider
    {
        private Dictionary<Validators, IValidateVBAIdentifiers> _nameOnlyValidators;

        private List<IEncapsulateFieldCandidate> _candidates;
        private List<IUserDefinedTypeMemberCandidate> _udtMemberCandidates;

        public EncapsulateFieldValidationsProvider()
        {
            _nameOnlyValidators = new Dictionary<Validators, IValidateVBAIdentifiers>()
            {
                [Validators.Default] = new IdentifierOnlyValidator(DeclarationType.Variable, false),
                [Validators.UserDefinedType] = new IdentifierOnlyValidator(DeclarationType.UserDefinedType, false),
                [Validators.UserDefinedTypeMember] = new IdentifierOnlyValidator(DeclarationType.UserDefinedTypeMember, false),
                [Validators.UserDefinedTypeMemberArray] = new IdentifierOnlyValidator(DeclarationType.UserDefinedTypeMember, true),
            };

            _candidates = new List<IEncapsulateFieldCandidate>();
            _udtMemberCandidates = new List<IUserDefinedTypeMemberCandidate>();
        }

        public IValidateVBAIdentifiers NameOnlyValidator(Validators validatorType)
            => _nameOnlyValidators[validatorType];

        public IEncapsulateFieldConflictFinder ConflictDetector(EncapsulateFieldStrategy strategy, IDeclarationFinderProvider declarationFinderProvider)
        {
            if (strategy == EncapsulateFieldStrategy.UseBackingFields)
            {
                return new ConflictDetectorUseBackingFields(declarationFinderProvider, _candidates, _udtMemberCandidates);
            }
            return new ConflictDetectorConvertFieldsToUDTMembers(declarationFinderProvider, _candidates, _udtMemberCandidates);
        }

        public void RegisterCandidate(IEncapsulateFieldCandidate candidate)
        {
            _candidates.Add(candidate);
        }

        public void RegisterCandidates(IEnumerable<IEncapsulateFieldCandidate> candidates)
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

        private class IdentifierOnlyValidator : IValidateVBAIdentifiers
        {
            private DeclarationType _declarationType;
            private bool _isArray;
            public IdentifierOnlyValidator(DeclarationType declarationType, bool isArray = false)
            {
                _declarationType = declarationType;
                _isArray = isArray;
            }

            public bool IsValidVBAIdentifier(string identifier, out string errorMessage)
                => !VBAIdentifierValidator.TryMatchInvalidIdentifierCriteria(identifier, _declarationType, out errorMessage, _isArray);

            public bool IsValidUDTMemberIdentifier(string identifier, bool isArray, out string errorMessage)
                => !VBAIdentifierValidator.TryMatchInvalidIdentifierCriteria(identifier, DeclarationType.UserDefinedTypeMember, out errorMessage, _isArray);
        }
    }
}
