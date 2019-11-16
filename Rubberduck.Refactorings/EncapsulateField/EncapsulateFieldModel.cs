using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.SmartIndenter;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface ISupportEncapsulateFieldTests
    {
        void SetMemberEncapsulationFlag(string name, bool flag);
        void SetEncapsulationFieldAttributes(string fieldName, IClientEditableFieldEncapsulationAttributes attributes);
    }
    public class EncapsulateFieldModel : IRefactoringModel, ISupportEncapsulateFieldTests
    {
        private readonly IIndenter _indenter;

        private IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> _udtFieldToUdtDeclarationMap = new Dictionary<Declaration, (Declaration, IEnumerable<Declaration>)>();
        private IEnumerable<Declaration> UdtFields => _udtFieldToUdtDeclarationMap.Keys;
        private IEnumerable<Declaration> UdtFieldMembers(Declaration udtField) => _udtFieldToUdtDeclarationMap[udtField].Item2;

        private Dictionary<KeyValuePair<Declaration, string>, IEncapsulatedFieldDeclaration> _encapsulateFieldDeclarations = new Dictionary<KeyValuePair<Declaration, string>, IEncapsulatedFieldDeclaration>();

        private IEncapsulatedFieldDeclaration _userSelectedEncapsulationField;
        private Dictionary<string, IFieldEncapsulationAttributes> _udtVariableEncapsulationAttributes = new Dictionary<string, IFieldEncapsulationAttributes>();

        public EncapsulateFieldModel(Declaration target, IEnumerable<Declaration> allMemberFields, IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> udtFieldToUdtDeclarationMap, IIndenter indenter)
        {
            _indenter = indenter;
            _udtFieldToUdtDeclarationMap = udtFieldToUdtDeclarationMap;

            foreach (var field in allMemberFields.Except(UdtFields))
            {
                var efd = EncapsulateDeclaration(field);
                _encapsulateFieldDeclarations.Add(efd.TargetIDPair, efd);
            }

            AddUDTEncapsulationFields(udtFieldToUdtDeclarationMap);

            var kvPair = _encapsulateFieldDeclarations.Where(efd => efd.Key.Key == target).Single();
            var selectedTarget = kvPair.Value; // _encapsulateFieldDeclarations[kvPair.Key];
            selectedTarget.EncapsulationAttributes.EncapsulateFlag = true;
            _userSelectedEncapsulationField = selectedTarget;
        }

        public IEnumerable<IEncapsulatedFieldDeclaration> FlaggedEncapsulationFields => _encapsulateFieldDeclarations.Values.Where(v => v.EncapsulationAttributes.EncapsulateFlag);

        public IEncapsulatedFieldDeclaration this[string encapsulatedFieldIdentifier]
        {
            get => _encapsulateFieldDeclarations.Values.Where(efd => efd.TargetIDPair.Value.Equals(encapsulatedFieldIdentifier))
                    .FirstOrDefault();
            set
            {
                var key = _encapsulateFieldDeclarations.Keys.Where(k => k.Value.Equals(encapsulatedFieldIdentifier))
                    .FirstOrDefault();
                _encapsulateFieldDeclarations[key] = value;
            }
        }

        public IList<string> PropertiesContent
        {
            get
            {
                var textBlocks = new List<string>();
                foreach (var field in FlaggedEncapsulationFields)
                {
                    textBlocks.Add(BuildPropertiesTextBlock(field.EncapsulationAttributes)); // as IFieldEncapsulationAttributes));
                }
                return textBlocks;
            }
        }

        private void AddUDTEncapsulationFields(IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> udtFieldToTypeMap)
        {
            foreach (var udtField in udtFieldToTypeMap.Keys)
            {
                var udtEncapsulation = DecorateUDTVariableDeclaration(udtField);
                _encapsulateFieldDeclarations.Add(udtEncapsulation.TargetIDPair, udtEncapsulation);


                foreach (var udtMember in UdtFieldMembers(udtField))
                {
                    var encapsulatedUdtMember = EncapsulateDeclaration(udtMember);
                    encapsulatedUdtMember = DecorateUDTMember(encapsulatedUdtMember, udtEncapsulation as EncapsulatedUserDefinedType);
                    _encapsulateFieldDeclarations.Add(encapsulatedUdtMember.TargetIDPair, encapsulatedUdtMember);
                }
            }
        }

        private IEncapsulatedFieldDeclaration EncapsulateDeclaration(Declaration target)
        {
            var encapsulated = new EncapsulatedFieldDeclaration(target);
            if (target.IsArray)
            {
                return EncapsulatedArrayType.Decorate(encapsulated);
            }
            else if (target.AsTypeName.Equals(Tokens.Variant))
            {
                return EncapsulatedVariantType.Decorate(encapsulated);
            }
            else if (target.IsObject)
            {
                return EncapsulatedObjectType.Decorate(encapsulated);
            }
            return EncapsulatedValueType.Decorate(encapsulated);
        }

        private IEncapsulatedFieldDeclaration DecorateUDTVariableDeclaration(Declaration target)
        {
            return EncapsulatedUserDefinedType.Decorate(new EncapsulatedFieldDeclaration(target));
        }

        private IEncapsulatedFieldDeclaration DecorateUDTMember(IEncapsulatedFieldDeclaration udtMember, EncapsulatedUserDefinedType udtVariable)
        {
            var targetIDPair = new KeyValuePair<Declaration, string>(udtMember.Declaration,$"{udtVariable.IdentifierName}.{udtMember.IdentifierName}");
            return EncapsulatedUserDefinedTypeMember.Decorate(udtMember, udtVariable, HasMultipleInstantiationsOfSameType(udtVariable.Declaration, targetIDPair));
        }

        private bool HasMultipleInstantiationsOfSameType(Declaration udtVariable, KeyValuePair<Declaration, string> targetIDPair)
        {
            var udt = _udtFieldToUdtDeclarationMap[udtVariable].Item1;
            var otherVariableOfTheSameType = _udtFieldToUdtDeclarationMap.Keys.Where(k => k != udtVariable && _udtFieldToUdtDeclarationMap[k].Item1 == udt);
            return otherVariableOfTheSameType.Any();
        }

        private string BuildPropertiesTextBlock(IFieldEncapsulationAttributes attributes)
        {
            var generator = new PropertyGenerator
            {
                PropertyName = attributes.PropertyName,
                AsTypeName = attributes.AsTypeName,
                BackingField = attributes.FieldReadWriteIdentifier,
                ParameterName = attributes.ParameterName,
                GenerateSetter = attributes.ImplementSetSetterType,
                GenerateLetter = attributes.ImplementLetSetterType
            };

            return GetPropertyText(generator);
        }

        private string GetPropertyText(PropertyGenerator generator)
        {
            var propertyTextLines = generator.AllPropertyCode.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            return string.Join(Environment.NewLine, _indenter.Indent(propertyTextLines, true));
        }

        public void SetMemberEncapsulationFlag(string name, bool flag)
        {
            this[name].EncapsulationAttributes.EncapsulateFlag = flag;
        }

        //Only used by tests....so far
        public void SetEncapsulationFieldAttributes(string fieldName, IClientEditableFieldEncapsulationAttributes attributes)
        {
            var currentAttributes = this[fieldName].EncapsulationAttributes; // as IClientEditableFieldEncapsulationAttributes;
            currentAttributes.NewFieldName = attributes.NewFieldName;
            currentAttributes.PropertyName = attributes.PropertyName;
            currentAttributes.EncapsulateFlag = attributes.EncapsulateFlag;
        }

        //This version only good for testing, fieldName could result in multiple results
        public void ApplyAttributes(string fieldName, IClientEditableFieldEncapsulationAttributes clientAttributes)
        {
            var encapsulatedField = this[fieldName];
            encapsulatedField.EncapsulationAttributes.NewFieldName = clientAttributes.NewFieldName;
            encapsulatedField.EncapsulationAttributes.PropertyName = clientAttributes.PropertyName;
            encapsulatedField.EncapsulationAttributes.ReadOnly = clientAttributes.ReadOnly;
            encapsulatedField.EncapsulationAttributes.EncapsulateFlag = clientAttributes.EncapsulateFlag;
        }

        public Declaration TargetDeclaration
        {
            get => _userSelectedEncapsulationField.Declaration;
            set
            {
                var encField = new EncapsulatedFieldDeclaration(value);
                _userSelectedEncapsulationField = _encapsulateFieldDeclarations[encField.TargetIDPair];
            }
        }

        public string PropertyName
        {
            get => _userSelectedEncapsulationField.EncapsulationAttributes.PropertyName;
            set => _userSelectedEncapsulationField.EncapsulationAttributes.PropertyName = value;
        }

        public string NewFieldName
        {
            get => _userSelectedEncapsulationField.EncapsulationAttributes.NewFieldName;
            set => _userSelectedEncapsulationField.EncapsulationAttributes.NewFieldName = value;
        }

        public string ParameterName
        {
            get => _userSelectedEncapsulationField.EncapsulationAttributes.ParameterName ?? "value";
            set => _userSelectedEncapsulationField.EncapsulationAttributes.ParameterName = value;
        }

        public bool IsReadOnly
        {
            get => _userSelectedEncapsulationField.EncapsulationAttributes.ReadOnly;
            set => _userSelectedEncapsulationField.EncapsulationAttributes.ReadOnly = value;
        }

        public bool EncapsulateFlag
        {
            get => _userSelectedEncapsulationField.EncapsulationAttributes.EncapsulateFlag;
            set => _userSelectedEncapsulationField.EncapsulationAttributes.EncapsulateFlag = value;
        }

        public bool ImplementLetSetterType
        {
            get => _userSelectedEncapsulationField.EncapsulationAttributes.ImplementLetSetterType;
            set => _userSelectedEncapsulationField.EncapsulationAttributes.ImplementLetSetterType = value;
        }

        public bool ImplementSetSetterType
        {
            get => _userSelectedEncapsulationField.EncapsulationAttributes.ImplementSetSetterType;
            set => _userSelectedEncapsulationField.EncapsulationAttributes.ImplementSetSetterType = value;
        }

        public bool CanImplementLet
            => _userSelectedEncapsulationField.EncapsulationAttributes.CanImplementLet; //.EncapsulationAttributes.IsValueType || _userSelectedEncapsulationField.EncapsulationAttributes.IsVariantType;

        public bool CanImplementSet
            => !_userSelectedEncapsulationField.EncapsulationAttributes.CanImplementSet; //IsValueType && _userSelectedEncapsulationField.EncapsulationAttributes.IsVariantType;
    }
}
