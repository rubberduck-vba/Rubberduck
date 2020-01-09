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
        bool HasConflictingIdentifier(IEncapsulatableField field, DeclarationType declarationType, out string errorMessage);
        IEncapsulatableField AssignNoConflictIdentifier(IEncapsulatableField candidate, DeclarationType declarationType);
        string CreateNonConflictIdentifierForProposedType(string identifier, QualifiedModuleName qmn, DeclarationType declarationType);
        bool IsConflictingProposedIdentifier(string fieldName, IEncapsulatableField candidate, DeclarationType declarationType);
        bool TryValidateEncapsulationAttributes(IEncapsulatableField field, out string errorMessage);
    }

    public class UseBackingFieldsConflictFinder : EncapsulateFieldConflictFinderBase
    {
        public UseBackingFieldsConflictFinder(IDeclarationFinderProvider declarationFinderProvider, IEnumerable<IEncapsulatableField> candidates, IEnumerable<IUserDefinedTypeMemberCandidate> udtCandidates)
            : base(declarationFinderProvider, candidates, udtCandidates) { }

        protected override IEnumerable<Declaration> FindRelevantMembers(IEncapsulatableField candidate)
            => _declarationFinderProvider.DeclarationFinder.Members(candidate.QualifiedModuleName)
                .Where(d => d != candidate.Declaration);
    }

    public class ConvertFieldsToUDTMembersConflictFinder : EncapsulateFieldConflictFinderBase
    {
        public ConvertFieldsToUDTMembersConflictFinder(IDeclarationFinderProvider declarationFinderProvider, IEnumerable<IEncapsulatableField> candidates, IEnumerable<IUserDefinedTypeMemberCandidate> udtCandidates)
            : base(declarationFinderProvider, candidates, udtCandidates) { }

        protected override IEnumerable<Declaration> FindRelevantMembers(IEncapsulatableField candidate)
        {
            var members = _declarationFinderProvider.DeclarationFinder.Members(candidate.QualifiedModuleName)
                .Where(d => d != candidate.Declaration);

            var membersToRemove = _fieldCandidates.Where(fc => fc.EncapsulateFlag && fc.Declaration.DeclarationType.HasFlag(DeclarationType.Variable))
                .Select(fc => fc.Declaration);

            return members.Except(membersToRemove);
        }
    }

    public abstract class EncapsulateFieldConflictFinderBase : IEncapsulateFieldConflictFinder
    {
        protected readonly IDeclarationFinderProvider _declarationFinderProvider;
        protected List<IEncapsulatableField> _fieldCandidates { set; get; } = new List<IEncapsulatableField>();
        protected List<IUserDefinedTypeMemberCandidate> _udtMemberCandidates { set; get; } = new List<IUserDefinedTypeMemberCandidate>();

        public EncapsulateFieldConflictFinderBase(IDeclarationFinderProvider declarationFinderProvider, IEnumerable<IEncapsulatableField> candidates, IEnumerable<IUserDefinedTypeMemberCandidate> udtCandidates)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _fieldCandidates.AddRange(candidates);
            _udtMemberCandidates.AddRange(udtCandidates);
        }

        public bool TryValidateEncapsulationAttributes(IEncapsulatableField field, out string errorMessage)
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

        public bool HasConflictingIdentifier(IEncapsulatableField field, DeclarationType declarationType, out string errorMessage)
            => InternalHasConflictingIdentifier(field, declarationType, false, out errorMessage);

        public IEncapsulatableField AssignNoConflictIdentifier(IEncapsulatableField candidate, DeclarationType declarationType)
        {
            var isConflictingIdentifier = HasConflictingIdentifierIgnoreEncapsulationFlag(candidate, declarationType, out _);
            var guard = 0;
            while (guard++ < 10 && isConflictingIdentifier)
            {
                var identifier = IdentifierToCompare(candidate, declarationType);

                if (declarationType.HasFlag(DeclarationType.Property))
                {
                    candidate.PropertyIdentifier = identifier.IncrementEncapsulationIdentifier();
                }
                else
                {
                    candidate.BackingIdentifier = identifier.IncrementEncapsulationIdentifier();
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

        public bool IsConflictingProposedIdentifier(string fieldName, IEncapsulatableField candidate, DeclarationType declarationType) 
            => PotentialConflictIdentifiers(candidate, declarationType)
                .Any(m => m.IsEquivalentVBAIdentifierTo(fieldName));

        protected abstract IEnumerable<Declaration> FindRelevantMembers(IEncapsulatableField candidate);

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

            var identifierToCompare = IdentifierToCompare(field, declarationType);

            if (potentialDeclarationIdentifierConflicts.Any(m => m.IsEquivalentVBAIdentifierTo(identifierToCompare)))
            {
                errorMessage = EncapsulateFieldResources.NameConflictDetected;
                return true;
            }
            return false;
        }

        protected string IdentifierToCompare(IEncapsulatableField field, DeclarationType declarationType)
        {
            return declarationType.HasFlag(DeclarationType.Property)
                ? field.PropertyIdentifier
                : field.BackingIdentifier;
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

        private List<IEncapsulatableField> FlaggedCandidates 
            => _fieldCandidates.Where(f => f.EncapsulateFlag).ToList();

        protected bool HasConflictingIdentifierIgnoreEncapsulationFlag(IEncapsulatableField field, DeclarationType declarationType, out string errorMessage)
            => InternalHasConflictingIdentifier(field, declarationType, true, out errorMessage);

        private bool IsConflictIdentifier(string fieldName, QualifiedModuleName qmn, DeclarationType declarationType)
        {
            var nameConflictCandidates = _declarationFinderProvider.DeclarationFinder.Members(qmn)
                .Where(d => !IsAlwaysIgnoreNameConflictType(d, declarationType)).ToList();

            return nameConflictCandidates.Any(m => m.IdentifierName.IsEquivalentVBAIdentifierTo(fieldName));
        }
    }
}
