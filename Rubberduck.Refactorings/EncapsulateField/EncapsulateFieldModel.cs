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
        void SetEncapsulationFieldAttributes(IUserModifiableFieldEncapsulationAttributes attributes);
    }
    public class EncapsulateFieldModel : IRefactoringModel, ISupportEncapsulateFieldTests
    {
        private readonly IIndenter _indenter;

        private IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> _udtFieldToUdtDeclarationMap = new Dictionary<Declaration, (Declaration, IEnumerable<Declaration>)>();
        private IEnumerable<Declaration> UdtFields => _udtFieldToUdtDeclarationMap.Keys;
        private IEnumerable<Declaration> UdtFieldMembers(Declaration udtField) => _udtFieldToUdtDeclarationMap[udtField].Item2;

        private Dictionary<Declaration, IEncapsulatedFieldDeclaration> _encapsulateFieldDeclarations = new Dictionary<Declaration, IEncapsulatedFieldDeclaration>();

        private IEncapsulatedFieldDeclaration _userSelectedEncapsulationField;
        private Dictionary<string, IFieldEncapsulationAttributes> _udtVariableEncapsulationAttributes = new Dictionary<string, IFieldEncapsulationAttributes>();

        public EncapsulateFieldModel(Declaration target, IEnumerable<Declaration> allMemberFields, IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> udtFieldToUdtDeclarationMap, IIndenter indenter)
        {
            _indenter = indenter;
            _udtFieldToUdtDeclarationMap = udtFieldToUdtDeclarationMap;

            foreach (var field in allMemberFields.Except(UdtFields))
            {
                AddEncapsulationField(EncapsulateDeclaration(field));
            }

            AddUDTEncapsulationFields(udtFieldToUdtDeclarationMap);

            this[target].EncapsulationAttributes.EncapsulateFlag = true;
            TargetDeclaration = target;
        }

        public IEnumerable<IEncapsulatedFieldDeclaration> FlaggedEncapsulationFields => _encapsulateFieldDeclarations.Values.Where(v => v.EncapsulationAttributes.EncapsulateFlag);

        public IEncapsulatedFieldDeclaration this[string encapsulatedFieldIdentifier]
        {
            get => _encapsulateFieldDeclarations.Values.Where(efd => efd.Declaration.IdentifierName.Equals(encapsulatedFieldIdentifier))
                    .FirstOrDefault();
            set
            {
                var key = _encapsulateFieldDeclarations.Keys.Where(k => k.IdentifierName.Equals(encapsulatedFieldIdentifier))
                    .FirstOrDefault();
                _encapsulateFieldDeclarations[key] = value;
            }
        }

        public IEncapsulatedFieldDeclaration this[Declaration declaration]
        {
            get
            {
                if (_encapsulateFieldDeclarations.TryGetValue(declaration, out var encapsulateFieldDefinition))
                {
                    return encapsulateFieldDefinition;
                }
                return null;
            }
            set
            {
                if (_encapsulateFieldDeclarations.ContainsKey(declaration))
                {
                    _encapsulateFieldDeclarations[declaration] = value;
                }
            }
        }

        public IList<string> PropertiesContent
        {
            get
            {
                var textBlocks = new List<string>();
                foreach (var field in FlaggedEncapsulationFields)
                {
                    textBlocks.Add(BuildPropertiesTextBlock(field.Declaration));
                }
                return textBlocks;
            }
        }

        private void AddUDTEncapsulationFields(IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> udtFieldToTypeMap)
        {
            foreach (var udtField in udtFieldToTypeMap.Keys)
            {
                var udtEncapsulation = DecorateUDTVariableDeclaration(udtField);
                AddEncapsulationField(udtEncapsulation);

                foreach (var udtMember in UdtFieldMembers(udtField))
                {
                    var efd = EncapsulateDeclaration(udtMember);
                    AddEncapsulationField(DecorateUDTMember(efd, udtEncapsulation as EncapsulatedUserDefinedType));
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
            return EncapsulatedUserDefinedTypeMember.Decorate(udtMember, udtVariable);
        }

        private void AddEncapsulationField(IEncapsulatedFieldDeclaration encapsulateFieldDeclaration)
        {
            if (_encapsulateFieldDeclarations.ContainsKey(encapsulateFieldDeclaration.Declaration))
            {
                _encapsulateFieldDeclarations[encapsulateFieldDeclaration.Declaration] = encapsulateFieldDeclaration;
                return;
            }
            _encapsulateFieldDeclarations.Add(encapsulateFieldDeclaration.Declaration, encapsulateFieldDeclaration);
        }

        private string BuildPropertiesTextBlock(Declaration target)
        {
            var attributes = this[target].EncapsulationAttributes;
            var generator = new PropertyGenerator
            {
                PropertyName = attributes.PropertyName,
                AsTypeName = attributes.AsTypeName,
                BackingField = attributes.NewFieldName,
                //BackingField = attributes.SetGet_LHSField,
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
        public void SetEncapsulationFieldAttributes(IUserModifiableFieldEncapsulationAttributes attributes)
        {
            var userAttributes = this[attributes.FieldName].EncapsulationAttributes as IUserModifiableFieldEncapsulationAttributes; // = attributes;
            userAttributes.FieldName = attributes.FieldName;
            userAttributes.NewFieldName = attributes.NewFieldName;
            userAttributes.PropertyName = attributes.PropertyName;
            userAttributes.EncapsulateFlag = attributes.EncapsulateFlag;
        }

        ////Only used by tests....so far
        //public void UpdateEncapsulationField(string variableName, string propertyName, bool encapsulateFlag, string newFieldName = null)
        //{
        //    var target = this[variableName];
        //    target.EncapsulationAttributes.FieldName = variableName;
        //    target.EncapsulationAttributes.PropertyName = propertyName;
        //    target.EncapsulationAttributes.NewFieldName = newFieldName ?? variableName;
        //    target.EncapsulationAttributes.EncapsulateFlag = encapsulateFlag;
        //    this[variableName] = target;
        //}

        //Only used by tests....so far
        public void UpdateEncapsulationField(IFieldEncapsulationAttributes attributes)
        {
            var target = this[attributes.FieldName];
            target.EncapsulationAttributes.FieldName = attributes.FieldName;
            target.EncapsulationAttributes.PropertyName = attributes.PropertyName;
            target.EncapsulationAttributes.NewFieldName = attributes.NewFieldName ?? attributes.FieldName;
            target.EncapsulationAttributes.EncapsulateFlag = attributes.EncapsulateFlag;
            this[attributes.FieldName] = target;
        }

        public Declaration TargetDeclaration
        {
            get => _userSelectedEncapsulationField.Declaration;
            set => _userSelectedEncapsulationField = _encapsulateFieldDeclarations[value];
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
