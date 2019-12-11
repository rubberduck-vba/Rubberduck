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
    public class EncapsulationCandidateFactory
    {
        private const string DEFAULT_STATE_UDT_IDENTIFIER = "This_Type";
        private const string DEFAULT_STATE_UDT_FIELD_IDENTIFIER = "this";

        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IEncapsulateFieldNamesValidator _validator;
        private List<IEncapsulateFieldCandidate> _encapsulatedFields = new List<IEncapsulateFieldCandidate>();
        //private IEnumerable<Declaration> _fieldCandidates;
        private QualifiedModuleName _targetQMN;

        public EncapsulationCandidateFactory(IDeclarationFinderProvider declarationFinderProvider, QualifiedModuleName targetQMN, IEncapsulateFieldNamesValidator validator)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _validator = validator;
            _targetQMN = targetQMN;
        }

        public IEncapsulateFieldCandidate CreateStateUDTField(string identifier = DEFAULT_STATE_UDT_FIELD_IDENTIFIER, string asTypeName = DEFAULT_STATE_UDT_IDENTIFIER)
        {
            var stateUDT = new StateUDTField(identifier, asTypeName, _targetQMN, _validator);
            _validator.RegisterFieldCandidate(stateUDT);

            stateUDT = SetNonConflictIdentifier(stateUDT, c => { return _validator.HasConflictingFieldIdentifier(stateUDT); }, (s) => { stateUDT.NewFieldName = s; }, () => stateUDT.IdentifierName, _validator);

            stateUDT = SetNonConflictIdentifier(stateUDT, c => { return _validator.IsConflictingStateUDTIdentifier(stateUDT); }, (s) => { stateUDT.AsTypeName = s; }, () => stateUDT.AsTypeName, _validator);
            return stateUDT;
        }

        public IEncapsulateFieldCandidate CreateEncapsulationCandidate(Declaration target)
        {
            Debug.Assert(!target.DeclarationType.Equals(DeclarationType.UserDefinedTypeMember));

            var candidate = target.IsUserDefinedTypeField()
                ? new EncapsulatedUserDefinedTypeField(target, _validator)
                : new EncapsulateFieldCandidate(target, _validator);

            _validator.RegisterFieldCandidate(candidate);

            candidate = ApplyTypeSpecificAttributes(candidate);
            candidate = SetNonConflictIdentifier(candidate, c => { return _validator.HasConflictingPropertyIdentifier(candidate); }, (s) => { candidate.PropertyName = s; }, () => candidate.IdentifierName, _validator);


            if (candidate is IEncapsulatedUserDefinedTypeField udtVariable)
            {
                (Declaration udt, IEnumerable<Declaration> udtMembers) = GetUDTAndMembersForField(udtVariable);

                udtVariable.TypeDeclarationIsPrivate = udt.Accessibility == Accessibility.Private;

                foreach (var udtMember in udtMembers)
                {
                    var candidateUDTMember = new EncapsulatedUserDefinedTypeMember(udtMember, udtVariable, _validator) as IEncapsulatedUserDefinedTypeMember;

                    candidateUDTMember = ApplyTypeSpecificAttributes(candidateUDTMember);

                    candidateUDTMember = SetNonConflictIdentifier(candidateUDTMember, c => { return _validator.HasConflictingPropertyIdentifier(candidate); }, (s) => { candidateUDTMember.PropertyName = s; }, () => candidate.IdentifierName, _validator);

                    candidateUDTMember.PropertyAccessExpression =
                       () =>
                       {
                           var prefix = udtVariable.EncapsulateFlag
                                      ? udtVariable.NewFieldName
                                      : udtVariable.IdentifierName;
                           return $"{prefix}.{candidateUDTMember.IdentifierName}";
                       };

                    udtVariable.AddMember(candidateUDTMember);
                }
            }
            return candidate;
        }

        public IEnumerable<IEncapsulateFieldCandidate> CreateEncapsulationCandidates() //IEnumerable<Declaration> candidateFields)
        {
            var fieldDeclarations = _declarationFinderProvider.DeclarationFinder
                .Members(_targetQMN)
                .Where(v => v.IsMemberVariable() && !v.IsWithEvents);
            //_fieldCandidates = candidateFields;

            var candidates = new List<IEncapsulateFieldCandidate>();
            foreach (var field in fieldDeclarations)
            {
                var fieldEncapsulationCandidate = CreateEncapsulationCandidate(field);

                candidates.Add(fieldEncapsulationCandidate);
                //_encapsulatedFields.Add(encapuslatedField);
            }

            return candidates;
        }

        private T SetNonConflictIdentifier<T>(T candidate, Predicate<T> conflictDetector, Action<string> setValue, Func<string> getIdentifier, IEncapsulateFieldNamesValidator validator) where T : IEncapsulateFieldCandidate
        {
            var isConflictingIdentifier = conflictDetector(candidate);
            for (var count = 1; count < 10 && isConflictingIdentifier; count++)
            {
                setValue(EncapsulationIdentifiers.IncrementIdentifier(getIdentifier()));
                isConflictingIdentifier = conflictDetector(candidate);
            }
            return candidate;
        }

        private T ApplyTypeSpecificAttributes<T>(T candidate) where T: IEncapsulateFieldCandidate
        {
            var target = candidate.Declaration;
            //if (target.IsUserDefinedTypeField())
            //{
            //    candidate.ImplementLetSetterType = true;
            //    candidate.ImplementSetSetterType = false;
            //}
            if (target.IsArray)
            {
                candidate.ImplementLetSetterType = false;
                candidate.ImplementSetSetterType = false;
                candidate.AsTypeName = Tokens.Variant;
                candidate.CanBeReadWrite = false;
                candidate.IsReadOnly = true;
            }
            else if (target.AsTypeName.Equals(Tokens.Variant))
            {
                candidate.ImplementLetSetterType = true;
                candidate.ImplementSetSetterType = true;
            }
            else if (target.IsObject)
            {
                candidate.ImplementLetSetterType = false;
                candidate.ImplementSetSetterType = true;
            }
            //else
            //{
            //    candidate.ImplementLetSetterType = true;
            //    candidate.ImplementSetSetterType = false;
            //}
            return candidate;
        }

        private (Declaration TypeDeclaration, IEnumerable<Declaration> Members) GetUDTAndMembersForField(IEncapsulatedUserDefinedTypeField udtField)
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
