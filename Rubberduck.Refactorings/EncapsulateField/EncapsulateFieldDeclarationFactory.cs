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

        public EncapsulationCandidateFactory(IDeclarationFinderProvider declarationFinderProvider, IEncapsulateFieldNamesValidator validator)
        {

            _declarationFinderProvider = declarationFinderProvider;
            _validator = validator;
        }

        public IEncapsulateFieldCandidate CreateStateUDTField(QualifiedModuleName qmn, string identifier = DEFAULT_STATE_UDT_FIELD_IDENTIFIER, string asTypeName = DEFAULT_STATE_UDT_IDENTIFIER)
        {
            var stateUDT = new StateUDTField(identifier, asTypeName, qmn, _validator);
            var isConflictingFieldIdentifier = _validator.HasConflictingFieldIdentifier(stateUDT);
            for (var count = 1; count < 10 && isConflictingFieldIdentifier; count++)
            {
                stateUDT.NewFieldName = EncapsulationIdentifiers.IncrementIdentifier(stateUDT.NewFieldName);
                isConflictingFieldIdentifier = _validator.HasConflictingFieldIdentifier(stateUDT);
            }

            var isConflictingUserDefinedTypeIdentifier = _validator.IsConflictingStateUDTIdentifier(stateUDT);
            for (var count = 1; count < 10 && isConflictingUserDefinedTypeIdentifier; count++)
            {
                stateUDT.AsTypeName = EncapsulationIdentifiers.IncrementIdentifier(stateUDT.AsTypeName);
                isConflictingUserDefinedTypeIdentifier = _validator.IsConflictingStateUDTIdentifier(stateUDT);
            }
            return stateUDT;
        }

        public IEnumerable<IEncapsulateFieldCandidate> CreateEncapsulationCandidates(IEnumerable<Declaration> candidateFields)
        {
            var candidates = new List<IEncapsulateFieldCandidate>();
            foreach (var field in candidateFields)
            {
                var encapuslatedField = EncapsulateDeclaration(field, _validator);

                var isConflictingPropertyIdentifier = _validator.HasConflictingPropertyIdentifier(encapuslatedField);
                for (var count = 1; count < 10 && isConflictingPropertyIdentifier; count++)
                {
                    encapuslatedField.PropertyName = $"{encapuslatedField.IdentifierName}{count}";
                    isConflictingPropertyIdentifier = _validator.HasConflictingPropertyIdentifier(encapuslatedField);
                }

                _encapsulatedFields.Add(encapuslatedField);
            }

            var udtFieldToUdtDeclarationMap = candidateFields
                .Where(v => v.IsUserDefinedTypeField())
                .Select(uv => CreateUDTTuple(uv))
                .ToDictionary(key => key.UDTVariable, element => (element.UserDefinedType, element.UDTMembers));

            foreach ( var udtField in udtFieldToUdtDeclarationMap.Keys)
            {
                var encapsulatedUDTField = _encapsulatedFields.Where(ef => ef.Declaration == udtField).Single() as IEncapsulatedUserDefinedTypeField;

                encapsulatedUDTField.TypeDeclarationIsPrivate = udtFieldToUdtDeclarationMap[udtField].UserDefinedType.Accessibility.Equals(Accessibility.Private);

                foreach (var udtMember in udtFieldToUdtDeclarationMap[udtField].Item2)
                {
                    var encapsulatedUDTMember = new EncapsulatedUserDefinedTypeMember(udtMember, encapsulatedUDTField, _validator) as IEncapsulatedUserDefinedTypeMember;

                    var isConflictingPropertyIdentifier = _validator.HasConflictingFieldIdentifier(encapsulatedUDTMember);
                    for (var count = 1; count < 10 && isConflictingPropertyIdentifier; count++)
                    {
                        encapsulatedUDTMember.PropertyName = $"{encapsulatedUDTMember.IdentifierName}{count}";
                        isConflictingPropertyIdentifier = _validator.HasConflictingPropertyIdentifier(encapsulatedUDTMember);
                    }

                    encapsulatedUDTMember = ApplyTypeSpecificAttributes(encapsulatedUDTMember);
                    encapsulatedUDTField.AddMember(encapsulatedUDTMember);
                    encapsulatedUDTMember.PropertyAccessExpression =
                       () =>
                       {
                           var prefix = encapsulatedUDTField.EncapsulateFlag
                                      ? encapsulatedUDTField.NewFieldName
                                      : encapsulatedUDTField.IdentifierName;
                            return $"{prefix}.{encapsulatedUDTMember.IdentifierName}";
                       };
                }
            }
            return _encapsulatedFields;
        }

        public static IEncapsulateFieldCandidate EncapsulateDeclaration(Declaration target, IEncapsulateFieldNamesValidator validator)
        {
            Debug.Assert(!target.DeclarationType.Equals(DeclarationType.UserDefinedTypeMember));

            var candidate = target.IsUserDefinedTypeField()
                ? new EncapsulatedUserDefinedTypeField(target, validator)
                : new EncapsulateFieldCandidate(target, validator);

            return ApplyTypeSpecificAttributes(candidate);
        }

        private (Declaration UDTVariable, Declaration UserDefinedType, IEnumerable<Declaration> UDTMembers) CreateUDTTuple(Declaration udtVariable)
        {
            var userDefinedTypeDeclaration = _declarationFinderProvider.DeclarationFinder
                .UserDeclarations(DeclarationType.UserDefinedType)
                .Where(ut => ut.IdentifierName.Equals(udtVariable.AsTypeName)
                    && (ut.Accessibility.Equals(Accessibility.Private)
                            && ut.QualifiedModuleName == udtVariable.QualifiedModuleName)
                    || (ut.Accessibility != Accessibility.Private))
                    .SingleOrDefault();

            var udtMembers = _declarationFinderProvider.DeclarationFinder
               .UserDeclarations(DeclarationType.UserDefinedTypeMember)
               .Where(utm => userDefinedTypeDeclaration == utm.ParentDeclaration);

            return (udtVariable, userDefinedTypeDeclaration, udtMembers);
        }

        private static T ApplyTypeSpecificAttributes<T>(T candidate) where T: IEncapsulateFieldCandidate
        {
            var target = candidate.Declaration;
            if (target.IsUserDefinedTypeField())
            {
                candidate.ImplementLetSetterType = true;
                candidate.ImplementSetSetterType = false;
            }
            else if (target.IsArray)
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
            else
            {
                candidate.ImplementLetSetterType = true;
                candidate.ImplementSetSetterType = false;
            }
            return candidate;
        }
    }
}
