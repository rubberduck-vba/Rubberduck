using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

        public IStateUDT CreateStateUDTField()
        {
            var stateUDT = new StateUDT(_targetQMN, _validator) as IStateUDT;

            stateUDT = SetNonConflictIdentifier(stateUDT, c => { return _validator.IsConflictingStateUDTFieldIdentifier(stateUDT); }, (s) => { stateUDT.FieldIdentifier = s; }, () => stateUDT.FieldIdentifier, _validator);

            stateUDT = SetNonConflictIdentifier(stateUDT, c => { return _validator.IsConflictingStateUDTTypeIdentifier(stateUDT); }, (s) => { stateUDT.TypeIdentifier = s; }, () => stateUDT.TypeIdentifier, _validator);

            return stateUDT;
        }

        public IEncapsulateFieldCandidate CreateEncapsulationCandidate(Declaration target)
        {
            Debug.Assert(!target.DeclarationType.Equals(DeclarationType.UserDefinedTypeMember));

            var candidate = target.IsUserDefinedTypeField()
                ? new UserDefinedTypeCandidate(target, _validator)
                : new EncapsulateFieldCandidate(target, _validator);

            _validator.RegisterFieldCandidate(candidate);

            candidate = ApplyTypeSpecificAttributes(candidate);

            candidate = SetNonConflictIdentifier(candidate, c => { return _validator.HasConflictingIdentifier(candidate, DeclarationType.Property); }, (s) => { candidate.PropertyName = s.Capitalize(); }, () => candidate.PropertyName, _validator);
            candidate = SetNonConflictIdentifier(candidate, c => { return _validator.HasConflictingIdentifier(candidate, DeclarationType.Variable); }, (s) => { candidate.FieldIdentifier = s.UnCapitalize(); }, () => candidate.FieldIdentifier, _validator);

            if (candidate is IUserDefinedTypeCandidate udtVariable)
            {
                (Declaration udtDeclaration, IEnumerable<Declaration> udtMembers) = GetUDTAndMembersForField(udtVariable);

                udtVariable.TypeDeclarationIsPrivate = udtDeclaration.HasPrivateAccessibility();

                foreach (var udtMember in udtMembers)
                {
                    var candidateUDTMember = new UserDefinedTypeMemberCandidate(udtMember, udtVariable, _validator) as IUserDefinedTypeMemberCandidate;

                    candidateUDTMember = ApplyTypeSpecificAttributes(candidateUDTMember);

                    candidateUDTMember = SetNonConflictIdentifier(candidateUDTMember, c => { return _validator.HasConflictingIdentifier(candidate, DeclarationType.Property); }, (s) => { candidateUDTMember.PropertyName = s.Capitalize(); }, () => candidate.IdentifierName, _validator);

                    udtVariable.AddMember(candidateUDTMember);
                }
            }
            return candidate;
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

        private T SetNonConflictIdentifier<T>(T candidate, Predicate<T> conflictDetector, Action<string> setValue, Func<string> getIdentifier, IEncapsulateFieldNamesValidator validator) where T : IEncapsulateFieldCandidate
        {
            var isConflictingIdentifier = conflictDetector(candidate);
            for (var count = 1; count < 10 && isConflictingIdentifier; count++)
            {
                setValue(getIdentifier().IncrementIdentifier());
                isConflictingIdentifier = conflictDetector(candidate);
            }
            return candidate;
        }

        private IStateUDT SetNonConflictIdentifier(IStateUDT candidate, Predicate<IStateUDT> conflictDetector, Action<string> setValue, Func<string> getIdentifier, IEncapsulateFieldNamesValidator validator)
        {
            var isConflictingIdentifier = conflictDetector(candidate);
            for (var count = 1; count < 10 && isConflictingIdentifier; count++)
            {
                setValue(getIdentifier().IncrementIdentifier());
                isConflictingIdentifier = conflictDetector(candidate);
            }
            return candidate;
        }

        private T ApplyTypeSpecificAttributes<T>(T candidate) where T: IEncapsulateFieldCandidate
        {
            //Default values are
            //candidate.ImplementLetSetterType = true;
            //candidate.ImplementSetSetterType = false;
            //candidate.CanBeReadWrite = true;
            //candidate.IsReadOnly = false;

            if (candidate.Declaration.IsArray)
            {
                candidate.ImplementLetSetterType = false;
                candidate.ImplementSetSetterType = false;
                candidate.AsTypeName = Tokens.Variant;
                candidate.CanBeReadWrite = false;
                candidate.IsReadOnly = true;
            }
            else if (candidate.Declaration.AsTypeName.Equals(Tokens.Variant))
            {
                candidate.ImplementLetSetterType = true;
                candidate.ImplementSetSetterType = true;
            }
            else if (candidate.Declaration.IsObject)
            {
                candidate.ImplementLetSetterType = false;
                candidate.ImplementSetSetterType = true;
            }
            return candidate;
        }

        private (Declaration TypeDeclaration, IEnumerable<Declaration> Members) GetUDTAndMembersForField(IUserDefinedTypeCandidate udtField)
        {
            var userDefinedTypeDeclaration = _declarationFinderProvider.DeclarationFinder
                .UserDeclarations(DeclarationType.UserDefinedType)
                .Where(ut => ut.IdentifierName.Equals(udtField.AsTypeName)
                    && (ut.Accessibility.Equals(Accessibility.Private)
                            && ut.QualifiedModuleName == udtField.QualifiedModuleName)
                    || (ut.Accessibility != Accessibility.Private))
                    .SingleOrDefault();

            var udtMembers = _declarationFinderProvider.DeclarationFinder
               .UserDeclarations(DeclarationType.UserDefinedTypeMember)
               .Where(utm => userDefinedTypeDeclaration == utm.ParentDeclaration);

            return (userDefinedTypeDeclaration, udtMembers);
        }
    }
}
