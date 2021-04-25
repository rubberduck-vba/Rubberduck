using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldConflictFinderFactory
    {
        IEncapsulateFieldConflictFinder Create(IDeclarationFinderProvider declarationFinderProvider,
            IEnumerable<IEncapsulateFieldCandidate> candidates,
            IEnumerable<IObjectStateUDT> objectStateUDTs);
    }

    public interface IEncapsulateFieldConflictFinder 
    {
        bool IsConflictingIdentifier(IEncapsulateFieldCandidate field, string identifierToCompare, out string errorMessage);
        (bool IsValid, string ValidationError) ValidateEncapsulationAttributes(IEncapsulateFieldCandidate field);
        void AssignNoConflictIdentifiers(IEncapsulateFieldCandidate candidate);
        void AssignNoConflictIdentifiers(IObjectStateUDT stateUDT);
        void AssignNoConflictIdentifiers(IEnumerable<IEncapsulateFieldCandidate> candidates);
        void AssignNoConflictBackingFieldIdentifier(IEncapsulateFieldCandidate candidate);
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

            _fieldCandidates.ForEach(c => LoadUDTMemberCandidates(c, _udtMemberCandidates));

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

            if (field.Declaration.IsArray)
            {
                if (field.Declaration.References.Any(rf => rf.QualifiedModuleName != field.QualifiedModuleName
                    && rf.Context.TryGetAncestor<VBAParser.RedimVariableDeclarationContext>(out _)))
                {
                    errorMessage = string.Format(RefactoringsUI.EncapsulateField_ArrayHasExternalRedimFormat, field.IdentifierName);
                    return (false, errorMessage);
                }

                if (field is IEncapsulateFieldAsUDTMemberCandidate udtMember
                    && VBAIdentifierValidator.TryMatchInvalidIdentifierCriteria(udtMember.UserDefinedTypeMemberIdentifier, declarationType, out errorMessage, true))
                {
                    return (false, errorMessage);
                }
            }

            var hasConflictFreeValidIdentifiers = 
                !VBAIdentifierValidator.TryMatchInvalidIdentifierCriteria(field.PropertyIdentifier, declarationType, out errorMessage, field.Declaration.IsArray)
                && !IsConflictingIdentifier(field, field.PropertyIdentifier, out errorMessage)
                && !IsConflictingIdentifier(field, field.BackingIdentifier, out errorMessage)
                && !(field is IEncapsulateFieldAsUDTMemberCandidate && ConflictsWithExistingUDTMembers(SelectedObjectStateUDT(), field.BackingIdentifier, out errorMessage));

            return (hasConflictFreeValidIdentifiers, errorMessage);
        }

        public bool IsConflictingIdentifier(IEncapsulateFieldCandidate field, string identifierToCompare, out string errorMessage)
        {
            errorMessage = string.Empty;
            if (HasConflictIdentifiers(field, identifierToCompare))
            {
                errorMessage = RefactoringsUI.EncapsulateField_NameConflictDetected;
            }
            return !string.IsNullOrEmpty(errorMessage);
        }

        public void AssignNoConflictIdentifiers(IEnumerable<IEncapsulateFieldCandidate> candidates)
        {
            foreach (var candidate in candidates.Where(c => c.EncapsulateFlag))
            {
                ResolveIdentifierConflicts(candidate);
            }
        }

        private void ResolveIdentifierConflicts(IEncapsulateFieldCandidate candidate)
        {
            AssignNoConflictIdentifiers(candidate);
            if (candidate is IUserDefinedTypeCandidate udtCandidate)
            {
                ResolveUDTMemberIdentifierConflicts(udtCandidate.Members);
            }
        }

        private void ResolveUDTMemberIdentifierConflicts(IEnumerable<IUserDefinedTypeMemberCandidate> members)
        {
            foreach (var member in members)
            {
                AssignNoConflictIdentifiers(member);
                if (member.WrappedCandidate is IUserDefinedTypeCandidate childUDT
                    && childUDT.Declaration.AsTypeDeclaration.HasPrivateAccessibility())
                {
                    ResolveIdentifierConflicts(childUDT);
                }
            }
        }

        public void AssignNoConflictIdentifiers(IEncapsulateFieldCandidate candidate)
        {
            if (candidate is IEncapsulateFieldAsUDTMemberCandidate udtMember)
            {
                AssignIdentifier(
                    () => ConflictsWithExistingUDTMembers(SelectedObjectStateUDT(), udtMember.UserDefinedTypeMemberIdentifier, out _),
                    () => udtMember.UserDefinedTypeMemberIdentifier = udtMember.UserDefinedTypeMemberIdentifier.IncrementEncapsulationIdentifier());
                return;
            }

            AssignNoConflictPropertyIdentifier(candidate);
            AssignNoConflictBackingFieldIdentifier(candidate);
        }

        public void AssignNoConflictIdentifiers(IObjectStateUDT stateUDT)
        {
            AssignIdentifier(
                () => _existingUserUDTsAndEnums.Any(m => m.IdentifierName.IsEquivalentVBAIdentifierTo(stateUDT.TypeIdentifier)),
                () => stateUDT.TypeIdentifier = stateUDT.TypeIdentifier.IncrementEncapsulationIdentifier());

            AssignIdentifier(
                () => HasConflictingFieldIdentifier(stateUDT, stateUDT.FieldIdentifier),
                () => stateUDT.FieldIdentifier = stateUDT.FieldIdentifier.IncrementEncapsulationIdentifier());
        }

        private IObjectStateUDT SelectedObjectStateUDT() 
            => _objectStateUDTs.SingleOrDefault(os => os.IsSelected);

        private static bool ConflictsWithExistingUDTMembers(IObjectStateUDT objectStateUDT, string identifier, out string errorMessage)
        {
            errorMessage = string.Empty;
            if (objectStateUDT?.ExistingMembers.Any(nm => nm.IdentifierName.IsEquivalentVBAIdentifierTo(identifier)) ?? false)
            {
                errorMessage = RefactoringsUI.EncapsulateField_NameConflictDetected;
            }
            return !string.IsNullOrEmpty(errorMessage);
        }

        private void AssignNoConflictPropertyIdentifier(IEncapsulateFieldCandidate candidate)
        {
            AssignIdentifier(
                () => IsConflictingIdentifier(candidate, candidate.PropertyIdentifier, out _),
                () => candidate.PropertyIdentifier = candidate.PropertyIdentifier.IncrementEncapsulationIdentifier());
        }

        public void AssignNoConflictBackingFieldIdentifier(IEncapsulateFieldCandidate candidate)
        {
            if (candidate.BackingIdentifierMutator != null)
            {
                AssignIdentifier(
                    () => IsConflictingIdentifier(candidate, candidate.BackingIdentifier, out _),
                    () => candidate.BackingIdentifierMutator(candidate.BackingIdentifier.IncrementEncapsulationIdentifier()));
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
            return HasInternalPropertyAndBackingFieldConflict(candidate)
                || HasConflictsWithOtherEncapsulationPropertyIdentifiers(candidate, identifierToCompare)
                || HasConflictsWithOtherEncapsulationBackingIdentifiers(candidate, identifierToCompare)
                || HasConflictsWithUnmodifiedPropertyAndFieldIdentifiers(candidate, identifierToCompare)
                || HasConflictWithLocalDeclarationIdentifiers(candidate, identifierToCompare);
        }

        private bool HasInternalPropertyAndBackingFieldConflict(IEncapsulateFieldCandidate candidate) 
            => candidate.BackingIdentifierMutator != null 
                && candidate.EncapsulateFlag
                && candidate.PropertyIdentifier.IsEquivalentVBAIdentifierTo(candidate.BackingIdentifier);

        private bool HasConflictsWithOtherEncapsulationPropertyIdentifiers(IEncapsulateFieldCandidate candidate, string identifierToCompare)
            => _allCandidates.Where(c => c.TargetID != candidate.TargetID
                && c.EncapsulateFlag
                && c.PropertyIdentifier.IsEquivalentVBAIdentifierTo(identifierToCompare)).Any();

        private bool HasConflictsWithOtherEncapsulationBackingIdentifiers(IEncapsulateFieldCandidate candidate, string identifierToCompare)
        {
            return candidate is IEncapsulateFieldAsUDTMemberCandidate || candidate is IUserDefinedTypeMemberCandidate
                ? false
                : _allCandidates.Where(c => c.TargetID != candidate.TargetID
                    && c.EncapsulateFlag
                    && c.BackingIdentifier.IsEquivalentVBAIdentifierTo(identifierToCompare)).Any();
        }

        private bool HasConflictsWithUnmodifiedPropertyAndFieldIdentifiers(IEncapsulateFieldCandidate candidate, string identifierToCompare)
        {
            var membersToEvaluate = _members.Where(d => d != candidate.Declaration);

            if (candidate is IEncapsulateFieldAsUDTMemberCandidate)
            {
                membersToEvaluate = membersToEvaluate.Except(
                    _fieldCandidates.Where(fc => fc.EncapsulateFlag && fc.Declaration.DeclarationType.HasFlag(DeclarationType.Variable))
                        .Select(f => f.Declaration));
            }

            var nameConflictCandidates = membersToEvaluate.Where(member => !(member.IsLocalVariable() || member.IsLocalConstant()
                || _declarationTypesThatNeverConflictWithFieldAndPropertyIdentifiers.Contains(member.DeclarationType)));

            return nameConflictCandidates.Any(m => m.IdentifierName.IsEquivalentVBAIdentifierTo(identifierToCompare));
        }

        private bool HasConflictWithLocalDeclarationIdentifiers(IEncapsulateFieldCandidate candidate, string identifierToCompare)
        {
            var membersToEvaluate = _members.Where(d => d != candidate.Declaration);

            if (candidate is IEncapsulateFieldAsUDTMemberCandidate)
            {
                membersToEvaluate = membersToEvaluate.Except(
                    _fieldCandidates.Where(fc => fc.EncapsulateFlag && fc.Declaration.DeclarationType.HasFlag(DeclarationType.Variable))
                        .Select(f => f.Declaration));
            }

            bool IsQualifedReference(IdentifierReference identifierReference)
                => identifierReference.Context.Parent is VBAParser.MemberAccessExprContext 
                    || identifierReference.Context.Parent is VBAParser.WithMemberAccessExprContext;

            //Only check IdentifierReferences in the declaring module because encapsulated field 
            //references in other modules will be module-qualified.
            var candidateLocalReferences = candidate.Declaration.References.Where(rf => rf.QualifiedModuleName == candidate.QualifiedModuleName
                && !IsQualifedReference(rf));

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

        private void LoadUDTMemberCandidates(IEncapsulateFieldCandidate candidate, List<IUserDefinedTypeMemberCandidate> udtMemberCandidates)
        {
            if (!(candidate is IUserDefinedTypeCandidate udtCandidate))
            {
                return;
            }

            foreach (var member in udtCandidate.Members)
            {
                udtMemberCandidates.Add(member);

                if (member.WrappedCandidate is IUserDefinedTypeCandidate childUDT
                    && childUDT.Declaration.AsTypeDeclaration.HasPrivateAccessibility())
                {
                    //recursive till a non-UserDefinedType member is found
                    LoadUDTMemberCandidates(childUDT, udtMemberCandidates);
                }
            }
        }
    }
}
