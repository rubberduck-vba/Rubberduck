using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Refactorings.Common;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldNamesValidator
    {
        void RegisterFieldCandidate(IEncapsulateFieldCandidate candidate);
        bool IsValidVBAIdentifier(string identifier, DeclarationType declarationType, out string errorMessage);
        bool HasConflictingIdentifier(IEncapsulateFieldCandidate candidate, DeclarationType declarationType, out string errorMessage);
        bool IsSelfConsistent(IEncapsulateFieldCandidate candidate, out string errorMessage);
        bool IsConflictingFieldIdentifier(string fieldName, IEncapsulateFieldCandidate candidate, DeclarationType declarationType);
        bool HasValidEncapsulationIdentifiers(IEncapsulateFieldCandidate candidate, out string errorMessage);
        bool IsConflictingStateUDTTypeIdentifier(IStateUDT stateUDT);
        bool IsConflictingStateUDTFieldIdentifier(IStateUDT stateUDT);
        IEncapsulateFieldCandidate AssignNoConflictIdentifier(IEncapsulateFieldCandidate candidate, DeclarationType declarationType);
    }

    public class EncapsulateFieldNamesValidator : IEncapsulateFieldNamesValidator
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        private List<IEncapsulateFieldCandidate> FieldCandidates { set; get; }

        private List<IUserDefinedTypeMemberCandidate> UDTMemberCandidates { set; get; }

        private List<IEncapsulateFieldCandidate> FlaggedCandidates => FieldCandidates.Where(rc => rc.EncapsulateFlag).ToList();

        public EncapsulateFieldNamesValidator(IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
            FieldCandidates = new List<IEncapsulateFieldCandidate>();
            UDTMemberCandidates = new List<IUserDefinedTypeMemberCandidate>();
        }

        public void RegisterFieldCandidate(IEncapsulateFieldCandidate candidate)
        {
            FieldCandidates.Add(candidate);
            if (candidate is IUserDefinedTypeCandidate udt)
            {
                foreach (var member in udt.Members)
                {
                    UDTMemberCandidates.Add(member);
                }
            }
        }

        private DeclarationFinder DeclarationFinder => _declarationFinderProvider.DeclarationFinder;


        public bool HasValidEncapsulationIdentifiers(IEncapsulateFieldCandidate candidate, out string errorMessage)
        {
            if (VBAIdentifierValidator.TryMatchInvalidIdentifierCriteria(candidate.PropertyName, DeclarationType.Property, out errorMessage))
            {
                return false;
            }
            if (!IsSelfConsistent(candidate, out errorMessage))
            {
                return false;
            }
            return true;
        }

        public bool IsValidVBAIdentifier(string identifier, DeclarationType declarationType, out string errorMessage)
        {
            return !VBAIdentifierValidator.TryMatchInvalidIdentifierCriteria(identifier, declarationType, out errorMessage);
        }

        public bool IsSelfConsistent(IEncapsulateFieldCandidate candidate, out string errorMessage)
        {
            errorMessage = string.Empty;
            if (candidate.PropertyName.IsEquivalentVBAIdentifierTo(candidate.FieldIdentifier))
            {
                errorMessage = $"{EncapsulateFieldResources.Conflict}: {EncapsulateFieldResources.Property}({candidate.PropertyName}) => {EncapsulateFieldResources.Field}({candidate.FieldIdentifier})";
                return false;
            }
            if (candidate.PropertyName.IsEquivalentVBAIdentifierTo(candidate.ParameterName))
            {
                errorMessage = $"{EncapsulateFieldResources.Conflict}: {EncapsulateFieldResources.Property}({candidate.PropertyName}) => {EncapsulateFieldResources.Parameter}({candidate.ParameterName})";
                return false;
            }
            if (candidate.FieldIdentifier.IsEquivalentVBAIdentifierTo(candidate.ParameterName))
            {
                errorMessage = $"{EncapsulateFieldResources.Conflict}: {EncapsulateFieldResources.Field}({candidate.FieldIdentifier}) => {EncapsulateFieldResources.Parameter}({candidate.ParameterName})";
                return false;
            }
            return true;
        }

        private List<string> PotentialConflictIdentifiers(IEncapsulateFieldCandidate candidate, DeclarationType declarationType)
        {
            var members = _declarationFinderProvider.DeclarationFinder.Members(candidate.QualifiedModuleName)
                .Where(d => d != candidate.Declaration);

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

        private List<string> PotentialConflictIdentifiers(IStateUDT stateUDT, DeclarationType declarationType)
        {
            var potentialDeclarationIdentifierConflicts = new List<string>();

            var members = _declarationFinderProvider.DeclarationFinder.Members(stateUDT.QualifiedModuleName);

            var nameConflictCandidates = members
                .Where(d => !IsAlwaysIgnoreNameConflictType(d, declarationType)).ToList();

            potentialDeclarationIdentifierConflicts.AddRange(nameConflictCandidates.Select(d => d.IdentifierName));

            return potentialDeclarationIdentifierConflicts.ToList();
        }

        public bool HasConflictingIdentifier(IEncapsulateFieldCandidate field, DeclarationType declarationType, out string errorMessage)
        {
            errorMessage = string.Empty;

            var potentialDeclarationIdentifierConflicts = new List<string>();
            potentialDeclarationIdentifierConflicts.AddRange(PotentialConflictIdentifiers(field, declarationType));

            potentialDeclarationIdentifierConflicts.AddRange(FlaggedCandidates.Where(fc => fc != field).Select(fc => fc.PropertyName));

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

        public IEncapsulateFieldCandidate AssignNoConflictIdentifier(IEncapsulateFieldCandidate candidate, DeclarationType declarationType) //, Predicate<IEncapsulateFieldCandidate> conflictDetector, Action<string> setValue, Func<string> getIdentifier)
        {
            var isConflictingIdentifier = HasConflictingIdentifier(candidate, declarationType, out _);
            for (var count = 1; count < 10 && isConflictingIdentifier; count++)
            {
                var identifier = declarationType.HasFlag(DeclarationType.Property)
                    ? candidate.PropertyName
                    : candidate.FieldIdentifier;

                if (declarationType.HasFlag(DeclarationType.Property))
                {
                    candidate.PropertyName = identifier.IncrementEncapsulationIdentifier();
                }
                else
                {
                    candidate.FieldIdentifier = identifier.IncrementEncapsulationIdentifier();
                }
                isConflictingIdentifier = HasConflictingIdentifier(candidate, declarationType, out _);
            }
            return candidate;
        }


        public bool IsConflictingFieldIdentifier(string fieldName, IEncapsulateFieldCandidate candidate, DeclarationType declarationType)
        {
            var members = PotentialConflictIdentifiers(candidate, declarationType);

            return members.Any(m => m.IsEquivalentVBAIdentifierTo(fieldName));
        }

        public bool IsConflictingStateUDTTypeIdentifier(IStateUDT stateUDT)
        {
            var potentialConflictNames = PotentialConflictIdentifiers(stateUDT, DeclarationType.UserDefinedType);

            return potentialConflictNames.Any(m => m.IsEquivalentVBAIdentifierTo(stateUDT.TypeIdentifier));
        }

        public bool IsConflictingStateUDTFieldIdentifier(IStateUDT stateUDT)
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
