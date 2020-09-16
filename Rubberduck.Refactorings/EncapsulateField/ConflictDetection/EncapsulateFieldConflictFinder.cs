using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
using Rubberduck.Resources;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldConflictFinder 
    {
        bool IsConflictingIdentifier(IEncapsulateFieldCandidate field, string identifierToCompare, out string errorMessage);
        (bool IsValid, string ValidationError) ValidateEncapsulationAttributes(IEncapsulateFieldCandidate field);
        void AssignNoConflictIdentifiers(IEncapsulateFieldCandidate candidate);
        void AssignNoConflictIdentifiers(IObjectStateUDT stateUDT);
        void AssignNoConflictIdentifiers(IEnumerable<IEncapsulateFieldCandidate> candidates);
    }

    public class EncapsulateFieldConflictFinder : IEncapsulateFieldConflictFinder
    {
        private static List<DeclarationType> _declarationTypesThatNeverConflictWithFieldAndPropertyIdentifiers = new List<DeclarationType>()
        {
            DeclarationType.Project,
            DeclarationType.ProceduralModule,
            DeclarationType.ClassModule,
            DeclarationType.Parameter,
            DeclarationType.Enumeration,
            DeclarationType.UserDefinedType,
            DeclarationType.UserDefinedTypeMember
        };

        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly List<IEncapsulateFieldCandidate> _fieldCandidates;
        private readonly List<IUserDefinedTypeMemberCandidate> _udtMemberCandidates;
        private readonly List<IEncapsulateFieldCandidate> _allCandidates;
        private readonly List<IObjectStateUDT> _objectStateUDTs;
        private readonly List<Declaration> _members;
        private readonly List<Declaration> _membersThatCanConflictWithFieldAndPropertyIdentifiers;
        private readonly List<Declaration> _existingUserUDTsAndEnums;

        public EncapsulateFieldConflictFinder(IDeclarationFinderProvider declarationFinderProvider, IEnumerable<IEncapsulateFieldCandidate> candidates, IEnumerable<IObjectStateUDT> objectStateUDTs)
        {
            _declarationFinderProvider = declarationFinderProvider;

            _fieldCandidates = candidates.ToList();

            _udtMemberCandidates = new List<IUserDefinedTypeMemberCandidate>();
            _fieldCandidates.ForEach(c => LoadUDTMembers(c));

            _allCandidates = _fieldCandidates.Concat(_udtMemberCandidates).ToList();

            _objectStateUDTs = objectStateUDTs.ToList();

            _members = _declarationFinderProvider.DeclarationFinder.Members(_allCandidates.First().QualifiedModuleName).ToList();

            _membersThatCanConflictWithFieldAndPropertyIdentifiers =
                _members.Where(m => !_declarationTypesThatNeverConflictWithFieldAndPropertyIdentifiers.Contains(m.DeclarationType)).ToList();

            _existingUserUDTsAndEnums = _members.Where(m => m.DeclarationType.HasFlag(DeclarationType.UserDefinedType)
                || m.DeclarationType.HasFlag(DeclarationType.Enumeration)).ToList();
        }

        public (bool IsValid, string ValidationError) ValidateEncapsulationAttributes(IEncapsulateFieldCandidate field)
        {
            if (!field.EncapsulateFlag)
            {
                return (true, string.Empty);
            }

            var declarationType = field is IEncapsulateFieldAsUDTMemberCandidate
                ? DeclarationType.UserDefinedTypeMember
                : field.Declaration.DeclarationType;

            var errorMessage = string.Empty;

            var hasInvalidIdentifierOrHasConflicts = 
                VBAIdentifierValidator.TryMatchInvalidIdentifierCriteria(field.PropertyIdentifier, declarationType, out errorMessage, field.Declaration.IsArray)
                || IsConflictingIdentifier(field, field.PropertyIdentifier, out errorMessage)
                || IsConflictingIdentifier(field, field.BackingIdentifier, out errorMessage)
                || field is IEncapsulateFieldAsUDTMemberCandidate && ConflictsWithExistingUDTMembers(SelectedObjectStateUDT(), field.BackingIdentifier, out errorMessage);

            return (string.IsNullOrEmpty(errorMessage), errorMessage);
        }

        public bool IsConflictingIdentifier(IEncapsulateFieldCandidate field, string identifierToCompare, out string errorMessage)
        {
            errorMessage = string.Empty;
            if (HasConflictIdentifiers(field, identifierToCompare))
            {
                errorMessage = RubberduckUI.EncapsulateField_NameConflictDetected;
            }
            return !string.IsNullOrEmpty(errorMessage);
        }

        public void AssignNoConflictIdentifiers(IEnumerable<IEncapsulateFieldCandidate> candidates)
        {
            foreach (var candidate in candidates.Where(c => c.EncapsulateFlag))
            {
                ResolveFieldConflicts(candidate);
            }
        }

        private void ResolveFieldConflicts(IEncapsulateFieldCandidate candidate)
        {
            AssignNoConflictIdentifiers(candidate);
            if (candidate is IUserDefinedTypeCandidate udtCandidate)
            {
                ResolveUDTMemberConflicts(udtCandidate.Members);
            }
        }

        private void ResolveUDTMemberConflicts(IEnumerable<IUserDefinedTypeMemberCandidate> members)
        {
            foreach (var member in members)
            {
                AssignNoConflictIdentifiers(member);
                if (member.WrappedCandidate is IUserDefinedTypeCandidate childUDT
                    && childUDT.Declaration.AsTypeDeclaration.HasPrivateAccessibility())
                {
                    ResolveFieldConflicts(childUDT);
                }
            }
        }

        public void AssignNoConflictIdentifiers(IEncapsulateFieldCandidate candidate)
        {
            if (candidate is IEncapsulateFieldAsUDTMemberCandidate)
            {
                AssignIdentifier(
                    () => ConflictsWithExistingUDTMembers(SelectedObjectStateUDT(), candidate.PropertyIdentifier, out _),
                    () => IncrementPropertyIdentifier(candidate));
                return;
            }

            AssignNoConflictPropertyIdentifier(candidate);
            AssignNoConflictBackingFieldIdentifier(candidate);
        }

        public void AssignNoConflictIdentifiers(IObjectStateUDT stateUDT)
        {
            AssignIdentifier(
                () => HasConflictingFieldIdentifier(stateUDT, stateUDT.FieldIdentifier),
                () => stateUDT.FieldIdentifier = stateUDT.FieldIdentifier.IncrementEncapsulationIdentifier());

            AssignIdentifier(
                () => _existingUserUDTsAndEnums.Any(m => m.IdentifierName.IsEquivalentVBAIdentifierTo(stateUDT.TypeIdentifier)), 
                () => stateUDT.TypeIdentifier = stateUDT.TypeIdentifier.IncrementEncapsulationIdentifier());
        }

        private IObjectStateUDT SelectedObjectStateUDT() => _objectStateUDTs.SingleOrDefault(os => os.IsSelected);

        private static bool ConflictsWithExistingUDTMembers(IObjectStateUDT objectStateUDT, string identifier, out string errorMessage)
        {
            errorMessage = string.Empty;
            if (objectStateUDT?.ExistingMembers.Any(nm => nm.IdentifierName.IsEquivalentVBAIdentifierTo(identifier)) ?? false)
            {
                errorMessage = RubberduckUI.EncapsulateField_NameConflictDetected;
            }
            return !string.IsNullOrEmpty(errorMessage);
        }

        private void IncrementPropertyIdentifier(IEncapsulateFieldCandidate candidate)
            => candidate.PropertyIdentifier = candidate.PropertyIdentifier.IncrementEncapsulationIdentifier();

        private void AssignNoConflictPropertyIdentifier(IEncapsulateFieldCandidate candidate)
        {
            AssignIdentifier(
                () => IsConflictingIdentifier(candidate, candidate.PropertyIdentifier, out _),
                () => IncrementPropertyIdentifier(candidate));
        }

        private void AssignNoConflictBackingFieldIdentifier(IEncapsulateFieldCandidate candidate)
        {
            //Private UserDefinedTypes are never used directly as a backing field - so never change their identifier.
            //The backing fields for an encapsulated Private UDT are its members.
            if (!(candidate is UserDefinedTypeMemberCandidate
                || candidate is IUserDefinedTypeCandidate udtCandidate && udtCandidate.TypeDeclarationIsPrivate))
            {
                AssignIdentifier(
                    () => IsConflictingIdentifier(candidate, candidate.BackingIdentifier, out _),
                    () => candidate.BackingIdentifier = candidate.BackingIdentifier.IncrementEncapsulationIdentifier());
            }
        }

        private static void AssignIdentifier(Func<bool> hasConflict, Action incrementIdentifier, int maxAttempts = 20)
        {
            var guard = 0;
            while (guard++ < maxAttempts && hasConflict())
            {
                incrementIdentifier();
            }

            if (guard >= maxAttempts)
            {
                throw new OverflowException("Unable to assign a non conflicting identifier");
            }
        }

        private bool HasConflictIdentifiers(IEncapsulateFieldCandidate candidate, string identifierToCompare)
        {
            if (_allCandidates.Where(c => c.TargetID != candidate.TargetID
                && c.EncapsulateFlag
                && c.PropertyIdentifier.IsEquivalentVBAIdentifierTo(identifierToCompare)).Any())
            {
                return true;
            }

            var membersToEvaluate = _members.Where(d => d != candidate.Declaration);

            if (candidate is IEncapsulateFieldAsUDTMemberCandidate)
            {
                membersToEvaluate = membersToEvaluate.Except(
                    _fieldCandidates.Where(fc => fc.EncapsulateFlag && fc.Declaration.DeclarationType.HasFlag(DeclarationType.Variable))
                        .Select(f => f.Declaration));
            }

            var nameConflictCandidates = membersToEvaluate.Where(member => !(member.IsLocalVariable() || member.IsLocalConstant()
                || _declarationTypesThatNeverConflictWithFieldAndPropertyIdentifiers.Contains(member.DeclarationType)));

            if (nameConflictCandidates.Any(m => m.IdentifierName.IsEquivalentVBAIdentifierTo(identifierToCompare)))
            {
                return true;
            }

            //Only check IdentifierReferences in the declaring module because IdentifierReferences in 
            //other modules will be module-qualified.
            var candidateLocalReferences = candidate.Declaration.References.Where(rf => rf.QualifiedModuleName == candidate.QualifiedModuleName);

            var localDeclarationConflictCandidates = membersToEvaluate.Where(localDec => candidateLocalReferences
               .Any(cr => cr.ParentScoping == localDec.ParentScopeDeclaration));

            return localDeclarationConflictCandidates.Any(m => m.IdentifierName.IsEquivalentVBAIdentifierTo(identifierToCompare));
        }

        private bool HasConflictingFieldIdentifier(IObjectStateUDT candidate, string identifierToCompare)
        {
            if (candidate.IsExistingDeclaration)
            {
                return false;
            }

            if (_allCandidates.Where(c => c.EncapsulateFlag
                && c.PropertyIdentifier.IsEquivalentVBAIdentifierTo(identifierToCompare)).Any())
            {
                return true;
            }

            var fieldsToRemoveFromConflictCandidates = _fieldCandidates
                .Where(fc => fc.EncapsulateFlag && fc.Declaration.DeclarationType.HasFlag(DeclarationType.Variable))
                .Select(fc => fc.Declaration);

            var nameConflictCandidates = 
                _members.Except(fieldsToRemoveFromConflictCandidates)
                    .Where(member => !(member.IsLocalVariable() || member.IsLocalConstant()
                        || _declarationTypesThatNeverConflictWithFieldAndPropertyIdentifiers.Contains(member.DeclarationType)));

            return nameConflictCandidates.Any(m => m.IdentifierName.IsEquivalentVBAIdentifierTo(identifierToCompare));
        }

        private void LoadUDTMembers(IEncapsulateFieldCandidate candidate)
        {
            if (!(candidate is IUserDefinedTypeCandidate udtCandidate))
            {
                return;
            }

            foreach (var member in udtCandidate.Members)
            {
                _udtMemberCandidates.Add(member);

                if (member.WrappedCandidate is IUserDefinedTypeCandidate childUDT
                    && childUDT.Declaration.AsTypeDeclaration.HasPrivateAccessibility())
                {
                    //recursive till a non-UserDefinedType member is found
                    LoadUDTMembers(childUDT);
                }
            }
        }
    }
}
