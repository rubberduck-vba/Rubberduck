using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
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

        public EncapsulateFieldElementsBuilder(IDeclarationFinderProvider declarationFinderProvider, QualifiedModuleName targetQMN)//, IEncapsulateFieldValidator validator )
        {
            _declarationFinderProvider = declarationFinderProvider;
            _targetQMN = targetQMN;
            CreateRefactoringElements();
        }

        public IObjectStateUDT ObjectStateUDT { private set; get; }

        public IEncapsulateFieldValidationsProvider ValidationsProvider { private set; get; }

        public IEnumerable<IEncapsulatableField> Candidates { private set; get; }

        private void CreateRefactoringElements()
        {
            var fieldDeclarations = _declarationFinderProvider.DeclarationFinder
                .Members(_targetQMN)
                .Where(v => v.IsMemberVariable() && !v.IsWithEvents);

            ValidationsProvider = new EncapsulateFieldValidationsProvider();

            var nameValidator = ValidationsProvider.NameOnlyValidator(NameValidators.Default);

            var candidates = new List<IEncapsulatableField>();
            foreach (var fieldDeclaration in fieldDeclarations)
            {
                Debug.Assert(!fieldDeclaration.DeclarationType.Equals(DeclarationType.UserDefinedTypeMember));

                var fieldEncapsulationCandidate = CreateCandidate(fieldDeclaration, nameValidator);


                candidates.Add(fieldEncapsulationCandidate);
            }

            ValidationsProvider.RegisterCandidates(candidates);

            var conflictsValidator = ValidationsProvider.ConflictDetector(EncapsulateFieldStrategy.UseBackingFields, _declarationFinderProvider);

            ObjectStateUDT = CreateStateUDTField(conflictsValidator);
            foreach (var candidate in candidates)
            {
                candidate.ConflictFinder = conflictsValidator;
                conflictsValidator.AssignNoConflictIdentifier(candidate, DeclarationType.Property);
                conflictsValidator.AssignNoConflictIdentifier(candidate, DeclarationType.Variable);
            }

            Candidates = candidates;
        }

        private IObjectStateUDT CreateStateUDTField(IEncapsulateFieldConflictFinder validator)
        {
            var stateUDT = new ObjectStateUDT(_targetQMN) as IObjectStateUDT;

            stateUDT.FieldIdentifier = validator.CreateNonConflictIdentifierForProposedType(stateUDT.FieldIdentifier, _targetQMN, DeclarationType.Variable);

            stateUDT.TypeIdentifier = validator.CreateNonConflictIdentifierForProposedType(stateUDT.TypeIdentifier, _targetQMN, DeclarationType.UserDefinedType);

            stateUDT.IsSelected = true;

            return stateUDT;
        }

        private IEncapsulatableField CreateCandidate(Declaration target, IValidateVBAIdentifiers validator)// Predicate<string> nameValidator)
        {
            if (target.IsUserDefinedTypeField())
            {
                var udtValidator = ValidationsProvider.NameOnlyValidator(NameValidators.UserDefinedType);
                var udtField = new UserDefinedTypeCandidate(target, udtValidator) as IUserDefinedTypeCandidate;

                (Declaration udtDeclaration, IEnumerable<Declaration> udtMembers) = GetUDTAndMembersForField(udtField);

                udtField.TypeDeclarationIsPrivate = udtDeclaration.HasPrivateAccessibility();

                udtField.NameValidator = udtValidator;

                foreach (var udtMemberDeclaration in udtMembers)
                {
                    var udtMemberValidator = ValidationsProvider.NameOnlyValidator(NameValidators.UserDefinedTypeMember);
                    if (udtMemberDeclaration.IsArray)
                    {
                        udtMemberValidator = ValidationsProvider.NameOnlyValidator(NameValidators.UserDefinedTypeMemberArray);
                    }
                    var candidateUDTMember = new UserDefinedTypeMemberCandidate(CreateCandidate(udtMemberDeclaration, udtMemberValidator), udtField) as IUserDefinedTypeMemberCandidate;

                    udtField.AddMember(candidateUDTMember);
                }

                var udtVariablesOfSameType = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                    .Where(v => v.AsTypeDeclaration == udtDeclaration);

                udtField.CanBeObjectStateUDT = udtField.TypeDeclarationIsPrivate && udtVariablesOfSameType.Count() == 1;

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
