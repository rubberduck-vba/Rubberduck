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
        private QualifiedModuleName _targetQMN;
        private readonly IEncapsulateFieldNamesValidator _validator;
        private IEnumerable<Declaration> _candidateFields;
        private bool _useNewStructure;

        public EncapsulationCandidateFactory(IDeclarationFinderProvider declarationFinderProvider, IEncapsulateFieldNamesValidator validator)
        {

            _declarationFinderProvider = declarationFinderProvider;
            _validator = validator;

            _useNewStructure = File.Exists("C:\\Users\\Brian\\Documents\\UseNewUDTStructure.txt");
        }

        public IEncapsulateFieldCandidate CreateProposedField(string identifier, string asTypeName, QualifiedModuleName qmn, IEncapsulateFieldNamesValidator validator)
        {
            var unselectableAttributes =  new UnselectableField(identifier, asTypeName, qmn, validator);
            return new EncapsulateFieldCandidate(unselectableAttributes, validator);
        }

        public IEnumerable<IEncapsulateFieldCandidate> CreateEncapsulationCandidates(IEnumerable<Declaration> candidateFields)
        {
            _candidateFields = candidateFields;

            _udtFieldToUdtDeclarationMap = candidateFields
                .Where(v => v.IsUserDefinedTypeField())
                .Select(uv => CreateUDTTuple(uv))
                .ToDictionary(key => key.UDTVariable, element => (element.UserDefinedType, element.UDTMembers));

            var candidates = new List<IEncapsulateFieldCandidate>();
            foreach (var field in candidateFields)
            {
                candidates.Add(EncapsulateDeclaration(field, _validator));
            }
            return candidates;
        }

        public IEncapsulateFieldCandidate EncapsulateDeclaration(Declaration target, IEncapsulateFieldNamesValidator validator)
        {
            if (target.IsUserDefinedTypeField())
            {
                var udtCandidate = new EncapsulatedUserDefinedTypeField(target, validator);
                udtCandidate.EncapsulationAttributes.ImplementLetSetterType = true;
                udtCandidate.EncapsulationAttributes.ImplementSetSetterType = false;

                if (_useNewStructure)
                {
                    var udtMembers = new List<EncapsulatedUserDefinedTypeMember>();

                    foreach (var member in _udtFieldToUdtDeclarationMap[target].Item2)
                    {
                        var udtMemberCandidate = EncapsulateDeclaration(member, validator);
                        var udtMember = new EncapsulatedUserDefinedTypeMember(udtMemberCandidate, udtCandidate, HasMultipleFieldsOfSameUserDefinedType(udtCandidate));
                        udtMembers.Add(udtMember);
                    }

                    udtCandidate.Members = udtMembers.Cast<IEncapsulateFieldCandidate>().ToList();
                }

                return udtCandidate;
            }

            var candidate = new EncapsulateFieldCandidate(target, validator);

            if (target.IsArray)
            {
                candidate.EncapsulationAttributes.ImplementLetSetterType = false;
                candidate.EncapsulationAttributes.ImplementSetSetterType = false;
                candidate.EncapsulationAttributes.AsTypeName = Tokens.Variant;
                candidate.CanBeReadWrite = false;
                candidate.IsReadOnly = true;
                return candidate;
            }
            else if (target.AsTypeName.Equals(Tokens.Variant))
            {
                candidate.EncapsulationAttributes.ImplementLetSetterType = true;
                candidate.EncapsulationAttributes.ImplementSetSetterType = true;
                return candidate;
            }
            else if (target.IsObject)
            {
                candidate.EncapsulationAttributes.ImplementLetSetterType = false;
                candidate.EncapsulationAttributes.ImplementSetSetterType = true;
                return candidate;
            }
            else if (target.DeclarationType.Equals(DeclarationType.UserDefinedTypeMember))
            {
                //TODO: This may need to pass back thru using it's AsTypeName
                candidate.EncapsulationAttributes.ImplementLetSetterType = true;
                candidate.EncapsulationAttributes.ImplementSetSetterType = false;
                return candidate;
            }

            candidate.EncapsulationAttributes.ImplementLetSetterType = true;
            candidate.EncapsulationAttributes.ImplementSetSetterType = false;
            return candidate;
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

        private bool HasMultipleFieldsOfSameUserDefinedType(IEncapsulateFieldCandidate udtCandidate)
            => _candidateFields.Any(fld => fld != udtCandidate.Declaration && udtCandidate.AsTypeName.Equals(fld.AsTypeName));
    }
}
