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
        private readonly IEncapsulateFieldNamesValidator _validator;
        private QualifiedModuleName _targetQMN;

        public EncapsulateFieldElementFactory(IDeclarationFinderProvider declarationFinderProvider, QualifiedModuleName targetQMN, IEncapsulateFieldNamesValidator validator)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _validator = validator;
            _targetQMN = targetQMN;
        }

        public IObjectStateUDT CreateStateUDTField()
        {
            var stateUDT = new ObjectStateUDT(_targetQMN) as IObjectStateUDT;

            stateUDT = SetNonConflictIdentifier(stateUDT, c => { return _validator.IsConflictingStateUDTFieldIdentifier(stateUDT); }, (s) => { stateUDT.FieldIdentifier = s; }, () => stateUDT.FieldIdentifier, _validator);

            stateUDT = SetNonConflictIdentifier(stateUDT, c => { return _validator.IsConflictingStateUDTTypeIdentifier(stateUDT); }, (s) => { stateUDT.TypeIdentifier = s; }, () => stateUDT.TypeIdentifier, _validator);

            return stateUDT;
        }

        public IEncapsulateFieldCandidate CreateEncapsulationCandidate(Declaration target)
        {
            Debug.Assert(!target.DeclarationType.Equals(DeclarationType.UserDefinedTypeMember));

            IEncapsulateFieldCandidate candidate = CreateCandidate(target);

            if (candidate is IUserDefinedTypeCandidate udtVariable)
            {
                (Declaration udtDeclaration, IEnumerable<Declaration> udtMembers) = GetUDTAndMembersForField(udtVariable);

                udtVariable.TypeDeclarationIsPrivate = udtDeclaration.HasPrivateAccessibility();

                foreach (var udtMemberDeclaration in udtMembers)
                {
                    var candidateUDTMember = new UserDefinedTypeMemberCandidate(CreateCandidate(udtMemberDeclaration), udtVariable, _validator) as IUserDefinedTypeMemberCandidate;

                    udtVariable.AddMember(candidateUDTMember);
                }

                var udtVariablesOfSameType = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                    .Where(v => v.AsTypeDeclaration == udtDeclaration);

                udtVariable.CanBeObjectStateUDT = udtVariable.TypeDeclarationIsPrivate && udtVariablesOfSameType.Count() == 1;
            }

            _validator.RegisterFieldCandidate(candidate);

            return candidate;
        }

        private IEncapsulateFieldCandidate CreateCandidate(Declaration target)
        {
            if (target.IsUserDefinedTypeField())
            {
                return new UserDefinedTypeCandidate(target, _validator);
            }
            else if (target.IsArray)
            {
                return new ArrayCandidate(target, _validator);
            }
            return new EncapsulateFieldCandidate(target, _validator);
        }

        public IEnumerable<IEncapsulateFieldCandidate> CreateEncapsulationCandidates()
        {
            var fieldDeclarations = _declarationFinderProvider.DeclarationFinder
                .Members(_targetQMN)
                .Where(v => v.IsMemberVariable() && !v.IsWithEvents);

            var candidates = new List<IEncapsulateFieldCandidate>();
            foreach (var field in fieldDeclarations)
            {
                var fieldEncapsulationCandidate = CreateEncapsulationCandidate(field);

                candidates.Add(fieldEncapsulationCandidate);
            }

            return candidates;
        }

        private IObjectStateUDT SetNonConflictIdentifier(IObjectStateUDT candidate, Predicate<IObjectStateUDT> conflictDetector, Action<string> setValue, Func<string> getIdentifier, IEncapsulateFieldNamesValidator validator)
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
