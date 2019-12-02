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
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> _udtFieldToUdtDeclarationMap = new Dictionary<Declaration, (Declaration, IEnumerable<Declaration>)>();
        private readonly IEncapsulateFieldNamesValidator _validator;
        private IEnumerable<Declaration> _candidateFields;
        private bool _useNewStructure;
        private List<IEncapsulateFieldCandidate> _encapsulatedFields = new List<IEncapsulateFieldCandidate>();

        public EncapsulationCandidateFactory(IDeclarationFinderProvider declarationFinderProvider, IEncapsulateFieldNamesValidator validator)
        {

            _declarationFinderProvider = declarationFinderProvider;
            _validator = validator;

            _useNewStructure = File.Exists("C:\\Users\\Brian\\Documents\\UseNewUDTStructure.txt");
        }

        public IEncapsulateFieldCandidate CreateInsertableField(string identifier, string asTypeName, QualifiedModuleName qmn, IEncapsulateFieldNamesValidator validator)
        {
            var unselectableAttributes =  new NeverEncapsulateAttributes(identifier, asTypeName, qmn, validator);
            return new EncapsulateFieldCandidate(unselectableAttributes, validator);
        }

        public IEnumerable<IEncapsulateFieldCandidate> CreateEncapsulationCandidates(IEnumerable<Declaration> candidateFields)
        {
            _candidateFields = candidateFields;

            var candidates = new List<IEncapsulateFieldCandidate>();
            foreach (var field in candidateFields)
            {
                _encapsulatedFields.Add(EncapsulateDeclaration(field, _validator));
            }

            _udtFieldToUdtDeclarationMap = candidateFields
                .Where(v => v.IsUserDefinedTypeField())
                .Select(uv => CreateUDTTuple(uv))
                .ToDictionary(key => key.UDTVariable, element => (element.UserDefinedType, element.UDTMembers));

            foreach ( var udtField in _udtFieldToUdtDeclarationMap.Keys)
            {
                var encapsulatedUDTField = _encapsulatedFields.Where(ef => ef.Declaration == udtField).Single();

                var moduleHasMultipleInstancesOfUDT = candidateFields.Any(fld => fld != encapsulatedUDTField.Declaration && encapsulatedUDTField.AsTypeName.Equals(fld.AsTypeName));
                var concreteParent = encapsulatedUDTField as EncapsulatedUserDefinedTypeField;

                foreach (var udtMember in _udtFieldToUdtDeclarationMap[udtField].Item2)
                {
                    IEncapsulateFieldCandidate encapsulatedUDTMember = new EncapsulatedUserDefinedTypeMember(udtMember, encapsulatedUDTField, _validator, moduleHasMultipleInstancesOfUDT);
                    encapsulatedUDTMember = ApplySpecializationAttributes(encapsulatedUDTMember);
                    concreteParent.Members.Add(encapsulatedUDTMember);
                }
            }
            return _encapsulatedFields;
        }

        private IEncapsulateFieldCandidate EncapsulateDeclaration(Declaration target, IEncapsulateFieldNamesValidator validator)
        {
            Debug.Assert(!target.DeclarationType.Equals(DeclarationType.UserDefinedTypeMember));

            var candidate = target.IsUserDefinedTypeField()
                ? new EncapsulatedUserDefinedTypeField(target, validator)
                : new EncapsulateFieldCandidate(target, validator);

            return ApplySpecializationAttributes(candidate);
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

        private IEncapsulateFieldCandidate ApplySpecializationAttributes(IEncapsulateFieldCandidate candidate)
        {
            var target = candidate.Declaration;
            if (target.IsUserDefinedTypeField())
            {
                candidate.EncapsulationAttributes.ImplementLetSetterType = true;
                candidate.EncapsulationAttributes.ImplementSetSetterType = false;
            }
            else if (target.IsArray)
            {
                candidate.EncapsulationAttributes.ImplementLetSetterType = false;
                candidate.EncapsulationAttributes.ImplementSetSetterType = false;
                candidate.EncapsulationAttributes.AsTypeName = Tokens.Variant;
                candidate.EncapsulationAttributes.CanBeReadWrite = false;
                candidate.EncapsulationAttributes.IsReadOnly = true;
            }
            else if (target.AsTypeName.Equals(Tokens.Variant))
            {
                candidate.EncapsulationAttributes.ImplementLetSetterType = true;
                candidate.EncapsulationAttributes.ImplementSetSetterType = true;
            }
            else if (target.IsObject)
            {
                candidate.EncapsulationAttributes.ImplementLetSetterType = false;
                candidate.EncapsulationAttributes.ImplementSetSetterType = true;
            }
            else
            {
                candidate.EncapsulationAttributes.ImplementLetSetterType = true;
                candidate.EncapsulationAttributes.ImplementSetSetterType = false;
            }
            return candidate;
        }
    }
}
