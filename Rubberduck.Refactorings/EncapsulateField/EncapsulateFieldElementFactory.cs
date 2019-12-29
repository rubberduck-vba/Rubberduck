using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldElementFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IEncapsulateFieldValidator _validator;
        private readonly IValidateEncapsulateFieldNames _namesValidator;
        private QualifiedModuleName _targetQMN;

        public EncapsulateFieldElementFactory(IDeclarationFinderProvider declarationFinderProvider, QualifiedModuleName targetQMN, IEncapsulateFieldValidator validator )
        {
            _declarationFinderProvider = declarationFinderProvider;
            _validator = validator;
            _namesValidator = validator as IValidateEncapsulateFieldNames;
            _targetQMN = targetQMN;
        }

        public IObjectStateUDT CreateStateUDTField()
        {
            var stateUDT = new ObjectStateUDT(_targetQMN) as IObjectStateUDT;

            stateUDT = SetNonConflictIdentifier(stateUDT, c => { return _validator.IsConflictingStateUDTFieldIdentifier(stateUDT); }, (s) => { stateUDT.FieldIdentifier = s; }, () => stateUDT.FieldIdentifier, _namesValidator);

            stateUDT = SetNonConflictIdentifier(stateUDT, c => { return _validator.IsConflictingStateUDTTypeIdentifier(stateUDT); }, (s) => { stateUDT.TypeIdentifier = s; }, () => stateUDT.TypeIdentifier, _namesValidator);

            return stateUDT;
        }

        private IEncapsulateFieldCandidate CreateCandidate(Declaration target)
        {
            if (target.IsUserDefinedTypeField())
            {
                var udtField = new UserDefinedTypeCandidate(target, _namesValidator) as IUserDefinedTypeCandidate;

                (Declaration udtDeclaration, IEnumerable<Declaration> udtMembers) = GetUDTAndMembersForField(udtField);

                udtField.TypeDeclarationIsPrivate = udtDeclaration.HasPrivateAccessibility();

                foreach (var udtMemberDeclaration in udtMembers)
                {
                    var candidateUDTMember = new UserDefinedTypeMemberCandidate(CreateCandidate(udtMemberDeclaration), udtField, _namesValidator) as IUserDefinedTypeMemberCandidate;

                    udtField.AddMember(candidateUDTMember);
                }

                var udtVariablesOfSameType = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                    .Where(v => v.AsTypeDeclaration == udtDeclaration);

                udtField.CanBeObjectStateUDT = udtField.TypeDeclarationIsPrivate && udtVariablesOfSameType.Count() == 1;

                return udtField;
            }
            else if (target.IsArray)
            {
                return new ArrayCandidate(target, _namesValidator);
            }
            return new EncapsulateFieldCandidate(target, _namesValidator);
        }

        public IEnumerable<IEncapsulateFieldCandidate> CreateEncapsulationCandidates()
        {
            var fieldDeclarations = _declarationFinderProvider.DeclarationFinder
                .Members(_targetQMN)
                .Where(v => v.IsMemberVariable() && !v.IsWithEvents);

            var candidates = new List<IEncapsulateFieldCandidate>();
            foreach (var fieldDeclaration in fieldDeclarations)
            {
                Debug.Assert(!fieldDeclaration.DeclarationType.Equals(DeclarationType.UserDefinedTypeMember));

                var fieldEncapsulationCandidate = CreateCandidate(fieldDeclaration);

                _validator.RegisterFieldCandidate(fieldEncapsulationCandidate);


                candidates.Add(fieldEncapsulationCandidate);
            }

            return candidates;
        }

        private IObjectStateUDT SetNonConflictIdentifier(IObjectStateUDT candidate, Predicate<IObjectStateUDT> conflictDetector, Action<string> setValue, Func<string> getIdentifier, IValidateEncapsulateFieldNames validator)
        {
            var isConflictingIdentifier = conflictDetector(candidate);
            for (var count = 1; count < 10 && isConflictingIdentifier; count++)
            {
                setValue(getIdentifier().IncrementEncapsulationIdentifier());
                isConflictingIdentifier = conflictDetector(candidate);
            }
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
