using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldConflictFinder 
    {
        //bool HasConflictingIdentifier(IEncapsulateFieldCandidate candidate, out string errorMessage);
        //bool HasConflictingIdentifier(IEncapsulateFieldCandidate field, DeclarationType declarationType, out string errorMessage);
        bool HasConflictingIdentifier(IEncapsulatableField candidate, out string errorMessage);
        bool HasConflictingIdentifier(IEncapsulatableField field, DeclarationType declarationType, out string errorMessage);
        IEncapsulateFieldCandidate AssignNoConflictIdentifier(IEncapsulateFieldCandidate candidate, DeclarationType declarationType);
        string CreateNonConflictIdentifierForProposedType(string identifier, QualifiedModuleName qmn, DeclarationType declarationType);
        bool IsConflictingProposedIdentifier(string fieldName, IEncapsulateFieldCandidate candidate, DeclarationType declarationType);
        bool TryValidateEncapsulationAttributes(IEncapsulatableField field, out string errorMessage);
    }

    public class ConflictDetectorUseBackingFields : ConflictDetectorBase
    {
        public ConflictDetectorUseBackingFields(IDeclarationFinderProvider declarationFinderProvider, IEnumerable<IEncapsulateFieldCandidate> candidates, IEnumerable<IUserDefinedTypeMemberCandidate> udtCandidates)
            : base(declarationFinderProvider, candidates, udtCandidates) { }

        public override bool TryValidateEncapsulationAttributes(IEncapsulatableField field, out string errorMessage)
        {
            errorMessage = string.Empty;
            if (!field.EncapsulateFlag) { return true; }

            if (!field.NameValidator.IsValidVBAIdentifier(field.PropertyIdentifier, out errorMessage))
            {
                return false;
            }

            if (HasConflictingIdentifier(field, DeclarationType.Property, out errorMessage))
            {
                return false;
            }
            return true;
        }

        protected override string IdentifierToCompare(IEncapsulatableField field, DeclarationType declarationType)
        {
            if (field is IUsingBackingField useBackingField)
            {
                return declarationType.HasFlag(DeclarationType.Property)
                    ? useBackingField.PropertyIdentifier
                    : useBackingField.FieldIdentifier;
            }
            return field.PropertyIdentifier;
        }

    }

    public class ConflictDetectorConvertFieldsToUDTMembers : ConflictDetectorBase
    {
        public ConflictDetectorConvertFieldsToUDTMembers(IDeclarationFinderProvider declarationFinderProvider, IEnumerable<IEncapsulateFieldCandidate> candidates, IEnumerable<IUserDefinedTypeMemberCandidate> udtCandidates)
            : base(declarationFinderProvider, candidates, udtCandidates) { }

        public override bool TryValidateEncapsulationAttributes(IEncapsulatableField field, out string errorMessage)
        {
            errorMessage = string.Empty;
            if (!field.EncapsulateFlag) { return true; }

            if (!field.NameValidator.IsValidVBAIdentifier(field.PropertyIdentifier, out errorMessage))
            {
                return false;
            }

            if (HasConflictingIdentifier(field,  DeclarationType.Property, out errorMessage))
            {
                return false;
            }

            return true;
        }

        protected override string IdentifierToCompare(IEncapsulatableField field, DeclarationType declarationType)
        {
            if (field is IConvertToUDTMember convertedField)
            {
                return declarationType.HasFlag(DeclarationType.Property)
                    ? convertedField.PropertyIdentifier
                    : convertedField.UDTMemberIdentifier;
            }
            return field.PropertyIdentifier;
        }

        //protected override IEnumerable<Declaration> FindRelevantMembers(IEncapsulateFieldCandidate candidate)
        //{
        //    var members = _declarationFinderProvider.DeclarationFinder.Members(candidate.QualifiedModuleName)
        //        .Where(d => d != candidate.Declaration);

        //    var membersToRemove = _fieldCandidates.Where(fc => fc.EncapsulateFlag && fc.Declaration.DeclarationType.HasFlag(DeclarationType.Variable))
        //        .Select(fc => fc.Declaration);

        //    return members.Except(membersToRemove);
        //}

        protected override IEnumerable<Declaration> FindRelevantMembers(IEncapsulatableField candidate)
        {
            var members = _declarationFinderProvider.DeclarationFinder.Members(candidate.QualifiedModuleName)
                .Where(d => d != candidate.Declaration);

            var membersToRemove = _fieldCandidates.Where(fc => fc.EncapsulateFlag && fc.Declaration.DeclarationType.HasFlag(DeclarationType.Variable))
                .Select(fc => fc.Declaration);

            return members.Except(membersToRemove);
        }
    }

    public abstract class ConflictDetectorBase : IEncapsulateFieldConflictFinder
    {
        protected readonly IDeclarationFinderProvider _declarationFinderProvider;
        protected List<IEncapsulateFieldCandidate> _fieldCandidates { set; get; } = new List<IEncapsulateFieldCandidate>();
        protected List<IUserDefinedTypeMemberCandidate> _udtMemberCandidates { set; get; } = new List<IUserDefinedTypeMemberCandidate>();

        public ConflictDetectorBase(IDeclarationFinderProvider declarationFinderProvider, IEnumerable<IEncapsulateFieldCandidate> candidates, IEnumerable<IUserDefinedTypeMemberCandidate> udtCandidates)//, IValidateVBAIdentifiers nameValidator)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _fieldCandidates.AddRange(candidates);
            _udtMemberCandidates.AddRange(udtCandidates);
        }

        public abstract bool TryValidateEncapsulationAttributes(IEncapsulatableField field, out string errorMessage);

        //public bool HasConflictingIdentifier(IEncapsulateFieldCandidate candidate, out string errorMessage)
        //    => InternalHasConflictingIdentifier(candidate, DeclarationType.Property, false, out errorMessage);

        public bool HasConflictingIdentifier(IEncapsulatableField candidate, out string errorMessage)
            => InternalHasConflictingIdentifier(candidate, DeclarationType.Property, false, out errorMessage);

        //public bool HasConflictingIdentifier(IEncapsulateFieldCandidate field, DeclarationType declarationType, out string errorMessage)
        //    => InternalHasConflictingIdentifier(field, declarationType, false, out errorMessage);

        public bool HasConflictingIdentifier(IEncapsulatableField field, DeclarationType declarationType, out string errorMessage)
            => InternalHasConflictingIdentifier(field, declarationType, false, out errorMessage);

        public IEncapsulateFieldCandidate AssignNoConflictIdentifier(IEncapsulateFieldCandidate candidate, DeclarationType declarationType) //, IValidateEncapsulateFieldNames validator)
        {
            var isConflictingIdentifier = HasConflictingIdentifierIgnoreEncapsulationFlag(candidate, declarationType, out _);
            var guard = 0;
            while (guard++ < 10 && isConflictingIdentifier)
            {
                var identifier = declarationType.HasFlag(DeclarationType.Property)
                    ? candidate.PropertyIdentifier
                    : candidate.FieldIdentifier;

                if (declarationType.HasFlag(DeclarationType.Property))
                {
                    candidate.PropertyIdentifier = identifier.IncrementEncapsulationIdentifier();
                }
                else
                {
                    candidate.FieldIdentifier = identifier.IncrementEncapsulationIdentifier();
                }
                isConflictingIdentifier = HasConflictingIdentifierIgnoreEncapsulationFlag(candidate, declarationType, out _);
            }
            return candidate;
        }

        public string CreateNonConflictIdentifierForProposedType(string identifier, QualifiedModuleName qmn, DeclarationType declarationType)
        {
            var guard = 0;
            while (guard++ < 10 && IsConflictIdentifier(identifier, qmn, declarationType))
            {
                identifier = identifier.IncrementEncapsulationIdentifier();
            }
            return identifier;
        }

        public bool IsConflictingProposedIdentifier(string fieldName, IEncapsulateFieldCandidate candidate, DeclarationType declarationType)
        {
            var members = PotentialConflictIdentifiers(candidate, declarationType);

            return members.Any(m => m.IsEquivalentVBAIdentifierTo(fieldName));
        }

        //protected virtual IEnumerable<Declaration> FindRelevantMembers(IEncapsulateFieldCandidate candidate)
        //    => _declarationFinderProvider.DeclarationFinder.Members(candidate.QualifiedModuleName)
        //        .Where(d => d != candidate.Declaration);

        protected virtual IEnumerable<Declaration> FindRelevantMembers(IEncapsulatableField candidate)
            => _declarationFinderProvider.DeclarationFinder.Members(candidate.QualifiedModuleName)
                .Where(d => d != candidate.Declaration);

        private bool InternalHasConflictingIdentifier(IEncapsulateFieldCandidate field, DeclarationType declarationType, bool ignoreEncapsulationFlags, out string errorMessage)
        {
            errorMessage = string.Empty;

            var potentialDeclarationIdentifierConflicts = new List<string>();
            potentialDeclarationIdentifierConflicts.AddRange(PotentialConflictIdentifiers(field, declarationType));

            if (ignoreEncapsulationFlags)
            {
                potentialDeclarationIdentifierConflicts.AddRange(_fieldCandidates.Where(fc => fc != field).Select(fc => fc.PropertyIdentifier));
            }
            else
            {
                potentialDeclarationIdentifierConflicts.AddRange(FlaggedCandidates.Where(fc => fc != field).Select(fc => fc.PropertyIdentifier));
            }

            potentialDeclarationIdentifierConflicts.AddRange(_udtMemberCandidates.Where(udtm => udtm != field && udtm.EncapsulateFlag).Select(udtm => udtm.PropertyIdentifier));

            var identifierToCompare = declarationType.HasFlag(DeclarationType.Property)
                ? field.PropertyIdentifier
                : field.FieldIdentifier;

            if (potentialDeclarationIdentifierConflicts.Any(m => m.IsEquivalentVBAIdentifierTo(identifierToCompare)))
            {
                errorMessage = EncapsulateFieldResources.NameConflictDetected;
                return true;
            }
            return false;
        }

        protected virtual bool InternalHasConflictingIdentifier(IEncapsulatableField field, DeclarationType declarationType, bool ignoreEncapsulationFlags, out string errorMessage)
        {
            errorMessage = string.Empty;

            var potentialDeclarationIdentifierConflicts = new List<string>();
            potentialDeclarationIdentifierConflicts.AddRange(PotentialConflictIdentifiers(field, declarationType));

            if (ignoreEncapsulationFlags)
            {
                potentialDeclarationIdentifierConflicts.AddRange(_fieldCandidates.Where(fc => fc != field).Select(fc => fc.PropertyIdentifier));
            }
            else
            {
                potentialDeclarationIdentifierConflicts.AddRange(FlaggedCandidates.Where(fc => fc != field).Select(fc => fc.PropertyIdentifier));
            }

            potentialDeclarationIdentifierConflicts.AddRange(_udtMemberCandidates.Where(udtm => udtm != field && udtm.EncapsulateFlag).Select(udtm => udtm.PropertyIdentifier));

            //var identifierToCompare = declarationType.HasFlag(DeclarationType.Property)
            //    ? field.PropertyIdentifier
            //    //: field.FieldIdentifier;
            //    : field.Declaration.IdentifierName; //TODO: Temporary
            var identifierToCompare = IdentifierToCompare(field, DeclarationType.Property);

            if (potentialDeclarationIdentifierConflicts.Any(m => m.IsEquivalentVBAIdentifierTo(identifierToCompare)))
            {
                errorMessage = EncapsulateFieldResources.NameConflictDetected;
                return true;
            }
            return false;
        }

        protected abstract string IdentifierToCompare(IEncapsulatableField field, DeclarationType declarationType);

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

        private List<string> PotentialConflictIdentifiers(IEncapsulateFieldCandidate candidate, DeclarationType declarationType)
        {
            var members = FindRelevantMembers(candidate);

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

        private List<string> PotentialConflictIdentifiers(IEncapsulatableField candidate, DeclarationType declarationType)
        {
            var members = FindRelevantMembers(candidate);

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

        private List<IEncapsulateFieldCandidate> FlaggedCandidates => _fieldCandidates.Where(f => f.EncapsulateFlag).ToList();

        private bool HasConflictingIdentifierIgnoreEncapsulationFlag(IEncapsulateFieldCandidate field, DeclarationType declarationType, out string errorMessage)
            => InternalHasConflictingIdentifier(field, declarationType, true, out errorMessage);

        private bool IsConflictIdentifier(string fieldName, QualifiedModuleName qmn, DeclarationType declarationType)
        {
            var nameConflictCandidates = _declarationFinderProvider.DeclarationFinder.Members(qmn)
                .Where(d => !IsAlwaysIgnoreNameConflictType(d, declarationType)).ToList();

            return nameConflictCandidates.Any(m => m.IdentifierName.IsEquivalentVBAIdentifierTo(fieldName));
        }
    }
}
