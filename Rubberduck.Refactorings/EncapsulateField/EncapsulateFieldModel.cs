using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.EncapsulateField.Strategies;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldModel : IRefactoringModel
    {
        private const string DEFAULT_ENCAPSULATION_UDT_IDENTIFIER = "This_Type";
        private const string DEFAULT_ENCAPSULATION_UDT_FIELD_IDENTIFIER = "this";

        private readonly IIndenter _indenter;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IEncapsulateFieldNamesValidator _validator;
        private readonly Func<EncapsulateFieldModel, string> _previewFunc;
        private QualifiedModuleName _targetQMN;

        private IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> _udtFieldToUdtDeclarationMap = new Dictionary<Declaration, (Declaration, IEnumerable<Declaration>)>();
        private IEnumerable<Declaration> UdtFields => _udtFieldToUdtDeclarationMap.Keys;
        private IEnumerable<Declaration> UdtFieldMembers(Declaration udtField) => _udtFieldToUdtDeclarationMap[udtField].Item2;

        private Dictionary<string, IEncapsulateFieldCandidate> _candidates = new Dictionary<string, IEncapsulateFieldCandidate>();

        public EncapsulateFieldModel(Declaration target, IDeclarationFinderProvider declarationFinderProvider, /*IEnumerable<Declaration> allMemberFields, IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> udtFieldToUdtDeclarationMap,*/ IIndenter indenter, IEncapsulateFieldNamesValidator validator, Func<EncapsulateFieldModel, string> previewFunc)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _indenter = indenter;
            _validator = validator;
            _previewFunc = previewFunc;
            _targetQMN = target.QualifiedModuleName;

            var encapsulationCandidateFields = _declarationFinderProvider.DeclarationFinder
                .Members(_targetQMN)
                .Where(v => v.IsMemberVariable() && !v.IsWithEvents);

            _udtFieldToUdtDeclarationMap = encapsulationCandidateFields
                .Where(v => v.AsTypeDeclaration?.DeclarationType.Equals(DeclarationType.UserDefinedType) ?? false)
                .Select(uv => CreateUDTTuple(uv))
                .ToDictionary(key => key.UDTVariable, element => (element.UserDefinedType, element.UDTMembers));

            foreach (var field in encapsulationCandidateFields.Except(UdtFields))
            {
                var efd = EncapsulateDeclaration(field);
                _candidates.Add(efd.TargetID, efd);
            }

            AddUDTEncapsulationFields(_udtFieldToUdtDeclarationMap);

            this[target].EncapsulationAttributes.EncapsulateFlag = true;

            //Maybe this should be passed in
            EncapsulationStrategy = new EncapsulateWithBackingFields(_targetQMN, _indenter);
        }

        public IEnumerable<IEncapsulateFieldCandidate> FlaggedEncapsulationFields 
            => _candidates.Values.Where(v => v.EncapsulationAttributes.EncapsulateFlag);

        public IEnumerable<IEncapsulateFieldCandidate> EncapsulationFields
            => _candidates.Values;

        public IEncapsulateFieldCandidate this[string encapsulatedFieldIdentifier] 
            => _candidates[encapsulatedFieldIdentifier];

        public IEncapsulateFieldCandidate this[Declaration fieldDeclaration] 
            => _candidates.Values.Where(efd => efd.Declaration.Equals(fieldDeclaration)).Select(encapsulatedField => encapsulatedField).Single();

        public IEncapsulateFieldStrategy EncapsulationStrategy { set; get; }

        public bool EncapsulateWithUDT
        {
            set
            {
                EncapsulationStrategy = value
                    ? new EncapsulateWithBackingUserDefinedType(_targetQMN, _indenter) as IEncapsulateFieldStrategy
                    : new EncapsulateWithBackingFields(_targetQMN, _indenter) as IEncapsulateFieldStrategy;
                //This probably should go - or be in the ctor
                EncapsulationStrategy.UdtMemberTargetIDToParentMap = _udtMemberTargetIDToParentMap;
            }

            get => EncapsulationStrategy is EncapsulateWithBackingUserDefinedType;
        }

        //What to do with these as far as IStrategy goes - only here to support the view model
        public string EncapsulateWithUDT_TypeIdentifier { set; get; } = DEFAULT_ENCAPSULATION_UDT_IDENTIFIER;

        public string EncapsulateWithUDT_FieldName { set; get; } = DEFAULT_ENCAPSULATION_UDT_FIELD_IDENTIFIER;

        public string PreviewRefactoring()
        {
            return _previewFunc(this);
        }

        private Dictionary<string, IEncapsulateFieldCandidate> _udtMemberTargetIDToParentMap { get; } = new Dictionary<string, IEncapsulateFieldCandidate>();
        private void AddUDTEncapsulationFields(IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> udtFieldToTypeMap)
        {
            foreach (var udtField in udtFieldToTypeMap.Keys)
            {
                var udtEncapsulation = EncapsulateDeclaration(udtField);
                _candidates.Add(udtEncapsulation.TargetID, udtEncapsulation);


                foreach (var udtMember in UdtFieldMembers(udtField))
                {
                    var encapsulatedUdtMember = EncapsulateDeclaration(udtMember);
                    encapsulatedUdtMember = DecorateUDTMember(encapsulatedUdtMember, udtEncapsulation as IEncapsulateFieldCandidate);
                    _candidates.Add(encapsulatedUdtMember.TargetID, encapsulatedUdtMember);
                    _udtMemberTargetIDToParentMap.Add(encapsulatedUdtMember.TargetID, udtEncapsulation);
                }
            }
        }

        private IEncapsulateFieldCandidate EncapsulateDeclaration(Declaration target) 
            => EncapsulateFieldDeclarationFactory.EncapsulateDeclaration(target, _validator);

        private IEncapsulateFieldCandidate DecorateUDTMember(IEncapsulateFieldCandidate udtMember, IEncapsulateFieldCandidate udtVariable)
        {
            var targetIDPair = new KeyValuePair<Declaration, string>(udtMember.Declaration,$"{udtVariable.Declaration.IdentifierName}.{udtMember.Declaration.IdentifierName}");
            return new EncapsulatedUserDefinedTypeMember(udtMember, udtVariable, HasMultipleInstantiationsOfSameType(udtVariable.Declaration, targetIDPair));
        }

        private bool HasMultipleInstantiationsOfSameType(Declaration udtVariable, KeyValuePair<Declaration, string> targetIDPair)
        {
            var udt = _udtFieldToUdtDeclarationMap[udtVariable].Item1;
            var otherVariableOfTheSameType = _udtFieldToUdtDeclarationMap.Keys.Where(k => k != udtVariable && _udtFieldToUdtDeclarationMap[k].Item1 == udt);
            return otherVariableOfTheSameType.Any();
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
               .Where(utm => userDefinedTypeDeclaration.IdentifierName == utm.ParentDeclaration.IdentifierName
                    && utm.QualifiedModuleName == userDefinedTypeDeclaration.QualifiedModuleName);

            return (udtVariable, userDefinedTypeDeclaration, udtMembers);
        }
    }
}
