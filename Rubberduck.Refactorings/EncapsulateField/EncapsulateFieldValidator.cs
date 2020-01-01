using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IObjectStateUDTNamesValidator
    {
        bool IsConflictingStateUDTTypeIdentifier(IObjectStateUDT stateUDT);
        bool IsConflictingStateUDTFieldIdentifier(IObjectStateUDT stateUDT);
    }

    public interface IValidateEncapsulateFieldNames
    {
        bool IsValidVBAIdentifier(string identifier, DeclarationType declarationType, out string errorMessage, bool isArray = false);
        bool HasConflictingIdentifier(IEncapsulateFieldCandidate candidate, DeclarationType declarationType, out string errorMessage);
        bool HasConflictingIdentifierIgnoreEncapsulationFlag(IEncapsulateFieldCandidate field, DeclarationType declarationType, out string errorMessage);
        bool IsConflictingProposedIdentifier(string fieldName, IEncapsulateFieldCandidate candidate, DeclarationType declarationType);
    }

    public interface IEncapsulateFieldValidator : IObjectStateUDTNamesValidator, IValidateEncapsulateFieldNames
    {
        void RegisterFieldCandidates(IEnumerable<IEncapsulateFieldCandidate> candidates);
    }

    public class EncapsulateFieldValidator : IEncapsulateFieldValidator
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        private List<IEncapsulateFieldCandidate> FieldCandidates { set; get; }

        private List<IUserDefinedTypeMemberCandidate> UDTMemberCandidates { set; get; }

        private List<IEncapsulateFieldCandidate> FlaggedCandidates => FieldCandidates.Where(rc => rc.EncapsulateFlag).ToList();

        public EncapsulateFieldValidator(IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
            FieldCandidates = new List<IEncapsulateFieldCandidate>();
            UDTMemberCandidates = new List<IUserDefinedTypeMemberCandidate>();
        }

        public void RegisterFieldCandidates(IEnumerable<IEncapsulateFieldCandidate> candidates)
        {
            FieldCandidates.AddRange(candidates);
            foreach (var udtCandidate in candidates.Where(c => c is IUserDefinedTypeCandidate).Cast<IUserDefinedTypeCandidate>())
            {
                foreach (var member in udtCandidate.Members)
                {
                    UDTMemberCandidates.Add(member);
                }
            }
        }

        private DeclarationFinder DeclarationFinder => _declarationFinderProvider.DeclarationFinder;

        public bool IsValidVBAIdentifier(string identifier, DeclarationType declarationType, out string errorMessage, bool isArray = false)
        {
            return !VBAIdentifierValidator.TryMatchInvalidIdentifierCriteria(identifier, declarationType, out errorMessage, isArray);
        }

        private List<string> PotentialConflictIdentifiers(IEncapsulateFieldCandidate candidate, DeclarationType declarationType)
        {
            var members = DeclarationFinder.Members(candidate.QualifiedModuleName)
                .Where(d => d != candidate.Declaration);

            if (candidate.ConvertFieldToUDTMember)
            {
                var membersToRemove = FieldCandidates.Where(fc => fc.EncapsulateFlag && fc.Declaration.DeclarationType.HasFlag(DeclarationType.Variable))
                    .Select(fc => fc.Declaration);

                members = members.Except(membersToRemove);
            }

            var nameConflictCandidates = members
                .Where(d => !IsAlwaysIgnoreNameConflictType(d, declarationType)).ToList();

            var localReferences = candidate.Declaration.References.Where(rf => rf.QualifiedModuleName == candidate.QualifiedModuleName);

            if (localReferences.Any())
            {
                foreach (var idRef in localReferences)
                {
                    var locals = members.Except(nameConflictCandidates)
                        .Where(localDec => localDec.ParentScopeDeclaration.Equals(idRef.ParentScoping));

                    nameConflictCandidates.AddRange(locals);
                }
            }
            return nameConflictCandidates.Select(c => c.IdentifierName).ToList();
        }

        private List<string> PotentialConflictIdentifiers(IObjectStateUDT stateUDT, DeclarationType declarationType)
        {
            var stateUDTDeclaration = stateUDT as IEncapsulateFieldDeclaration;
            var potentialDeclarationIdentifierConflicts = new List<string>();

            var members = DeclarationFinder.Members(stateUDTDeclaration.QualifiedModuleName);

            var nameConflictCandidates = members
                .Where(d => !IsAlwaysIgnoreNameConflictType(d, declarationType)).ToList();

            potentialDeclarationIdentifierConflicts.AddRange(nameConflictCandidates.Select(d => d.IdentifierName));

            return potentialDeclarationIdentifierConflicts.ToList();
        }

        public bool HasConflictingIdentifierIgnoreEncapsulationFlag(IEncapsulateFieldCandidate field, DeclarationType declarationType, out string errorMessage)
            => InternalHasConflictingIdentifier(field, declarationType, true, out errorMessage);

        public bool HasConflictingIdentifier(IEncapsulateFieldCandidate field, DeclarationType declarationType, out string errorMessage)
            => InternalHasConflictingIdentifier(field, declarationType, false, out errorMessage);

        private bool InternalHasConflictingIdentifier(IEncapsulateFieldCandidate field, DeclarationType declarationType, bool ignoreEncapsulationFlags, out string errorMessage)
        {
            errorMessage = string.Empty;

            var potentialDeclarationIdentifierConflicts = new List<string>();
            potentialDeclarationIdentifierConflicts.AddRange(PotentialConflictIdentifiers(field, declarationType));

            if (ignoreEncapsulationFlags)
            {
                potentialDeclarationIdentifierConflicts.AddRange(FieldCandidates.Where(fc => fc != field).Select(fc => fc.PropertyName));
            }
            else
            {
                potentialDeclarationIdentifierConflicts.AddRange(FlaggedCandidates.Where(fc => fc != field).Select(fc => fc.PropertyName));
            }
            potentialDeclarationIdentifierConflicts.AddRange(UDTMemberCandidates.Where(udtm => udtm != field && udtm.EncapsulateFlag).Select(udtm => udtm.PropertyName));

            var identifierToCompare = declarationType.HasFlag(DeclarationType.Property)
                ? field.PropertyName
                : field.FieldIdentifier;

            if (potentialDeclarationIdentifierConflicts.Any(m => m.IsEquivalentVBAIdentifierTo(identifierToCompare)))
            {
                errorMessage = EncapsulateFieldResources.NameConflictDetected;
                return true;
            }
            return false;
        }

        public bool IsConflictingProposedIdentifier(string fieldName, IEncapsulateFieldCandidate candidate, DeclarationType declarationType)
        {
            var members = PotentialConflictIdentifiers(candidate, declarationType);

            return members.Any(m => m.IsEquivalentVBAIdentifierTo(fieldName));
        }

        public bool IsConflictingStateUDTTypeIdentifier(IObjectStateUDT stateUDT)
        {
            var potentialConflictNames = PotentialConflictIdentifiers(stateUDT, DeclarationType.UserDefinedType);

            return potentialConflictNames.Any(m => m.IsEquivalentVBAIdentifierTo(stateUDT.TypeIdentifier));
        }

        public bool IsConflictingStateUDTFieldIdentifier(IObjectStateUDT stateUDT)
        {
            var potentialConflictNames = PotentialConflictIdentifiers(stateUDT, DeclarationType.Variable);

            return potentialConflictNames.Any(m => m.IsEquivalentVBAIdentifierTo(stateUDT.FieldIdentifier));
        }

       //The refactoring only inserts new code elements with the following Accessibilities:
       //Variables => Private
       //Properties => Public
       //UDT => Private
        private bool IsAlwaysIgnoreNameConflictType(Declaration d, DeclarationType toEnapsulateDeclarationType)
        {
            var NeverCauseNameConflictTypes = new List<DeclarationType>()
            {
                DeclarationType.Project,
                DeclarationType.ProceduralModule,
                DeclarationType.ClassModule,
                DeclarationType.Parameter,
                DeclarationType.EnumerationMember,
                DeclarationType.Enumeration,
                DeclarationType.UserDefinedType,
                DeclarationType.UserDefinedTypeMember
            };

            if (toEnapsulateDeclarationType.HasFlag(DeclarationType.Variable))
            {
                //5.2.3.4: An enum member name may not be the same as any variable name
                //or constant name that is defined within the same module
                NeverCauseNameConflictTypes.Remove(DeclarationType.EnumerationMember);
            }
            else if (toEnapsulateDeclarationType.HasFlag(DeclarationType.UserDefinedType))
            {
                //5.2.3.3 If an<udt-declaration > is an element of a<private-type-declaration> its 
                //UDT name cannot be the same as the enum name of any<enum-declaration> 
                //or the UDT name of any other<UDTdeclaration> within the same<module>
                NeverCauseNameConflictTypes.Remove(DeclarationType.UserDefinedType);

                //5.2.3.4 The enum name of a <private-enum-declaration> cannot be the same as the enum name of any other 
                //<enum-declaration> or as the UDT name of a <UDT-declaration> within the same <module>.
                NeverCauseNameConflictTypes.Remove(DeclarationType.Enumeration);
            }
            else if (toEnapsulateDeclarationType.HasFlag(DeclarationType.Property))
            {
                //Each < subroutine - declaration > and < function - declaration > must have a 
                //procedure name that is different from any other module variable name, 
                //module constant name, enum member name, or procedure name that is defined 
                //within the same module.

                NeverCauseNameConflictTypes.Remove(DeclarationType.EnumerationMember);
            }
            return d.IsLocalVariable()
                    || d.IsLocalConstant()
                    || NeverCauseNameConflictTypes.Contains(d.DeclarationType);
        }
    }
}
