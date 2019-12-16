using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Refactorings.Common;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldValidator
    {
        void RegisterFieldCandidate(IEncapsulateFieldCandidate candidate);
        bool IsValidVBAIdentifier(string identifier, DeclarationType declarationType);
        bool HasConflictingIdentifier(IEncapsulateFieldCandidate candidate, DeclarationType declarationType);
        bool IsConflictingFieldIdentifier(string fieldName, IEncapsulateFieldCandidate candidate, DeclarationType declarationType);
    }

    public interface IEncapsulateFieldNamesValidator : IEncapsulateFieldValidator
    {
        bool IsConflictingStateUDTTypeIdentifier(IStateUDT stateUDT);
        bool IsConflictingStateUDTFieldIdentifier(IStateUDT stateUDT);
    }

    public class EncapsulateFieldNamesValidator : IEncapsulateFieldNamesValidator
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        private List<IEncapsulateFieldCandidate> FieldCandidates { set; get; }

        private List<IEncapsulateFieldCandidate> FlaggedCandidates => FieldCandidates.Where(rc => rc.EncapsulateFlag).ToList();

        public EncapsulateFieldNamesValidator(IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
            FieldCandidates = new List<IEncapsulateFieldCandidate>();
        }

        public void RegisterFieldCandidate(IEncapsulateFieldCandidate candidate)
        {
            FieldCandidates.Add(candidate);
        }

        private DeclarationFinder DeclarationFinder => _declarationFinderProvider.DeclarationFinder;

        public bool IsValidVBAIdentifier(string identifier, DeclarationType declarationType) 
            => VBAIdentifierValidator.IsValidIdentifier(identifier, declarationType);

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

        private List<string> GetPotentialConflictMembers(IStateUDT stateUDT, DeclarationType declarationType)
        {
            var potentialDeclarationIdentifierConflicts = new List<string>();

            var members = _declarationFinderProvider.DeclarationFinder.Members(stateUDT.QualifiedModuleName);

            var nameConflictCandidates = members
                .Where(d => !IsAlwaysIgnoreNameConflictType(d, declarationType)).ToList();

            potentialDeclarationIdentifierConflicts.AddRange(nameConflictCandidates.Select(d => d.IdentifierName));

            return potentialDeclarationIdentifierConflicts.ToList();
        }

        public bool HasConflictingIdentifier(IEncapsulateFieldCandidate field, DeclarationType declarationType)
        {
            //The declared type of a function declaration may not be a private enum name.
            //=>Means encapsulating a private enum field must be 'As Long'

            //If a procedure declaration whose visibility is public has a procedure name 
            //that is the same as the name of a project or name of a module then 
            //all references to the procedure name must be explicitly qualified with 
            //its project or module name unless the reference occurs within the module that defines the procedure. 


         var potentialDeclarationIdentifierConflicts = new List<string>();
            potentialDeclarationIdentifierConflicts.AddRange(PotentialConflictIdentifiers(field, declarationType));

            potentialDeclarationIdentifierConflicts.AddRange(FlaggedCandidates.Where(fc => fc.Declaration != field.Declaration).Select(fc => fc.PropertyName));
            //potentialDeclarationIdentifierConflicts.AddRange(FlaggedCandidates.Where(fc => fc.Declaration != field.Declaration).Select(fc => fc.FieldIdentifier));

            var identifierToCompare = declarationType.HasFlag(DeclarationType.Property)
                ? field.PropertyName
                : field.FieldIdentifier;

            return potentialDeclarationIdentifierConflicts.Any(m => m.EqualsVBAIdentifier(identifierToCompare));
        }

        public bool IsConflictingFieldIdentifier(string fieldName, IEncapsulateFieldCandidate candidate, DeclarationType declarationType)
        {
            var members = PotentialConflictIdentifiers(candidate, declarationType);

            return members.Any(m => m.EqualsVBAIdentifier(fieldName));
        }

        public bool IsConflictingStateUDTTypeIdentifier(IStateUDT stateUDT)
        {
            var potentialConflictNames = GetPotentialConflictMembers(stateUDT, DeclarationType.UserDefinedType);

            return potentialConflictNames.Any(m => m.EqualsVBAIdentifier(stateUDT.TypeIdentifier));
        }

        public bool IsConflictingStateUDTFieldIdentifier(IStateUDT stateUDT)
        {
            var potentialConflictNames = GetPotentialConflictMembers(stateUDT, DeclarationType.Variable);

            return potentialConflictNames.Any(m => m.EqualsVBAIdentifier(stateUDT.FieldIdentifier));
        }

       //The refactoring only inserts elements with the following Accessibilities:
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
