using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldElementsBuilder
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private QualifiedModuleName _targetQMN;
        private string _defaultObjectStateUDTTypeName;
        private ICodeBuilder _codeBuilder;

        public EncapsulateFieldElementsBuilder(IDeclarationFinderProvider declarationFinderProvider, QualifiedModuleName targetQMN)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _targetQMN = targetQMN;
            _defaultObjectStateUDTTypeName = $"T{_targetQMN.ComponentName}";
            _codeBuilder = new CodeBuilder();
            CreateRefactoringElements();
        }

        public IObjectStateUDT DefaultObjectStateUDT { private set; get; }

        public IObjectStateUDT ObjectStateUDT { private set; get; }

        public IEncapsulateFieldValidationsProvider ValidationsProvider { private set; get; }

        public IEnumerable<IEncapsulateFieldCandidate> Candidates { private set; get; }

        public IEnumerable<IObjectStateUDT> ObjectStateUDTCandidates { private set; get; } = new List<IObjectStateUDT>();

        private void CreateRefactoringElements()
        {
            var fieldDeclarations = _declarationFinderProvider.DeclarationFinder
                .Members(_targetQMN)
                .Where(v => v.IsMemberVariable() && !v.IsWithEvents);

            var defaultNamesValidator = EncapsulateFieldValidationsProvider.NameOnlyValidator(NameValidators.Default);

            var candidates = new List<IEncapsulateFieldCandidate>();
            foreach (var fieldDeclaration in fieldDeclarations)
            {
                Debug.Assert(!fieldDeclaration.DeclarationType.Equals(DeclarationType.UserDefinedTypeMember));

                var fieldEncapsulationCandidate = CreateCandidate(fieldDeclaration, defaultNamesValidator);

                candidates.Add(fieldEncapsulationCandidate);
            }

            Candidates = candidates;

            ObjectStateUDTCandidates = BuildObjectStateUDTCandidates(candidates).ToList();

            ObjectStateUDT = ObjectStateUDTCandidates.FirstOrDefault(os => os.AsTypeDeclaration.IdentifierName.StartsWith(_defaultObjectStateUDTTypeName, StringComparison.InvariantCultureIgnoreCase));

            DefaultObjectStateUDT = CreateStateUDTField();
            DefaultObjectStateUDT.IsSelected = true;
            if (ObjectStateUDT != null)
            {
                ObjectStateUDT.IsSelected = true;
                DefaultObjectStateUDT.IsSelected = false;
            }

            ObjectStateUDTCandidates = ObjectStateUDTCandidates.Concat(new IObjectStateUDT[] { DefaultObjectStateUDT });

            ValidationsProvider = new EncapsulateFieldValidationsProvider(Candidates, ObjectStateUDTCandidates);

            var conflictsFinder = ValidationsProvider.ConflictDetector(EncapsulateFieldStrategy.UseBackingFields, _declarationFinderProvider);
            foreach (var candidate in candidates)
            {
                candidate.ConflictFinder = conflictsFinder;
            }
        }

        private IEnumerable<IObjectStateUDT> BuildObjectStateUDTCandidates(IEnumerable<IEncapsulateFieldCandidate> candidates)
        {
            var udtCandidates = candidates.Where(c => c is IUserDefinedTypeCandidate udt
                        && udt.CanBeObjectStateUDT);

            var objectStateUDTs = new List<IObjectStateUDT>();
            foreach (var udt in udtCandidates)
            {
                objectStateUDTs.Add(new ObjectStateUDT(udt as IUserDefinedTypeCandidate));
            }

            var objectStateUDT = objectStateUDTs.FirstOrDefault(os => os.AsTypeDeclaration.IdentifierName.StartsWith(_defaultObjectStateUDTTypeName, StringComparison.InvariantCultureIgnoreCase));

            return objectStateUDTs;
        }

        private IObjectStateUDT CreateStateUDTField()
        {
            var stateUDT = new ObjectStateUDT(_targetQMN) as IObjectStateUDT;

            EncapsulateFieldValidationsProvider.AssignNoConflictIdentifiers(stateUDT, _declarationFinderProvider);

            stateUDT.IsSelected = true;

            return stateUDT;
        }

        private IEncapsulateFieldCandidate CreateCandidate(Declaration target, IValidateVBAIdentifiers validator)
        {
            if (target.IsUserDefinedType())
            {
                var udtValidator = EncapsulateFieldValidationsProvider.NameOnlyValidator(NameValidators.UserDefinedType);
                var udtField = new UserDefinedTypeCandidate(target, udtValidator) as IUserDefinedTypeCandidate;

                (Declaration udtDeclaration, IEnumerable<Declaration> udtMembers) = GetUDTAndMembersForField(udtField);

                udtField.TypeDeclarationIsPrivate = udtDeclaration.HasPrivateAccessibility();

                udtField.NameValidator = udtValidator;

                foreach (var udtMemberDeclaration in udtMembers)
                {
                    var udtMemberValidator = EncapsulateFieldValidationsProvider.NameOnlyValidator(NameValidators.UserDefinedTypeMember);
                    if (udtMemberDeclaration.IsArray)
                    {
                        udtMemberValidator = EncapsulateFieldValidationsProvider.NameOnlyValidator(NameValidators.UserDefinedTypeMemberArray);
                    }
                    var candidateUDTMember = new UserDefinedTypeMemberCandidate(CreateCandidate(udtMemberDeclaration, udtMemberValidator), udtField) as IUserDefinedTypeMemberCandidate;

                    udtField.AddMember(candidateUDTMember);
                }

                var udtVariablesOfSameType = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                    .Where(v => v.AsTypeDeclaration == udtDeclaration);

                udtField.CanBeObjectStateUDT = udtField.TypeDeclarationIsPrivate 
                    && udtField.Declaration.HasPrivateAccessibility()
                    && udtVariablesOfSameType.Count() == 1;

                return udtField;
            }
            else if (target.IsArray)
            {
                return new ArrayCandidate(target, validator);
            }

            var candidate = new EncapsulateFieldCandidate(target, validator);
            return candidate;
        }

        private (Declaration TypeDeclaration, IEnumerable<Declaration> Members) GetUDTAndMembersForField(IUserDefinedTypeCandidate udtField)
        {
            var userDefinedTypeDeclaration = udtField.Declaration.AsTypeDeclaration;

            var udtMembers = _declarationFinderProvider.DeclarationFinder
               .UserDeclarations(DeclarationType.UserDefinedTypeMember)
               .Where(utm => userDefinedTypeDeclaration == utm.ParentDeclaration);

            return (userDefinedTypeDeclaration, udtMembers);
        }
    }
}
