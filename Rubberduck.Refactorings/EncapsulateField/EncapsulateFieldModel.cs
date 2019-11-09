using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldModel : IRefactoringModel
    {
        private readonly IIndenter _indenter;

        private IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> _udtFieldToUdtDeclarationMap = new Dictionary<Declaration, (Declaration, IEnumerable<Declaration>)>();
        private IEnumerable<Declaration> UdtFields => _udtFieldToUdtDeclarationMap.Keys;
        private IEnumerable<Declaration> UdtFieldMembers(Declaration udtField) => _udtFieldToUdtDeclarationMap[udtField].Item2;

        private Dictionary<Declaration, EncapsulateFieldDeclaration> _encapsulateFieldDeclarations = new Dictionary<Declaration, EncapsulateFieldDeclaration>();

        private EncapsulateFieldDeclaration _userSelectedEncapsulationField;
        private Dictionary<string, IUDTFieldEncapsulationAttributes> _udtVariableEncapsulationAttributes = new Dictionary<string, IUDTFieldEncapsulationAttributes>();

        public EncapsulateFieldModel(Declaration target, IEnumerable<Declaration> allMemberFields, IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> udtFieldToUdtDeclarationMap, IIndenter indenter)
        {
            _indenter = indenter;
            _udtFieldToUdtDeclarationMap = udtFieldToUdtDeclarationMap;

            foreach (var field in allMemberFields.Except(UdtFields))
            {
                AddEncapsulationField(DecorateDeclaration(field));
            }

            AddUDTEncapsulationFields(udtFieldToUdtDeclarationMap);

            this[target].EncapsulationAttributes.IsFlaggedToEncapsulate = true;
            TargetDeclaration = target;
            //_userSelectedEncapsulationField = this[target];
        }

        public IEnumerable<EncapsulateFieldDeclaration> FlaggedEncapsulationFields => _encapsulateFieldDeclarations.Values.Where(v => v.EncapsulationAttributes.IsFlaggedToEncapsulate);

        private EncapsulateFieldDeclaration DecorateDeclaration(Declaration target)
        {
            //TODO: Array type fields
            var selectedField = new EncapsulateFieldDeclaration(target);

            if (selectedField.EncapsulationAttributes.IsVariantType)
            {
                return new EncapsulateVariantType(selectedField);
            }
            else if (selectedField.EncapsulationAttributes.IsObjectType)
            {
                return new EncapsulateObjectType(selectedField);
            }
            else if (selectedField.EncapsulationAttributes.IsArray)
            {
                return new EncapsulateArrayType(selectedField);
            }
            return new EncapsulateValueType(selectedField);
        }

        private EncapsulateFieldDeclaration DecorateUDTVariableDeclaration(Declaration target)
        {
            return new EncapsulateUserDefinedType(new EncapsulateFieldDeclaration(target));
        }

        private EncapsulateFieldDeclaration DecorateUDTMember(Declaration udtMember, EncapsulateUserDefinedType udtVariable)
        {
            var selectedField = new EncapsulateFieldDeclaration(udtMember);
            if (selectedField.EncapsulationAttributes.IsVariantType)
            {
                return new EncapsulatedUserDefinedMemberValueType(new EncapsulateVariantType(selectedField), udtVariable);
            }
            else if (selectedField.EncapsulationAttributes.IsObjectType)
            {
                return new EncapsulatedUserDefinedMemberObjectType(new EncapsulateObjectType(selectedField), udtVariable);
            }
            return new EncapsulatedUserDefinedMemberValueType(new EncapsulateValueType(selectedField), udtVariable);
        }

        public void AddUDTVariableAttributes(IUDTFieldEncapsulationAttributes udtFieldEncapsulationAttributes)
        {
            if (_udtVariableEncapsulationAttributes.TryGetValue(udtFieldEncapsulationAttributes.FieldName, out _))
            {
                _udtVariableEncapsulationAttributes[udtFieldEncapsulationAttributes.FieldName] = udtFieldEncapsulationAttributes;
                return;
            }
            _udtVariableEncapsulationAttributes.Add(udtFieldEncapsulationAttributes.FieldName, udtFieldEncapsulationAttributes);
        }

        public bool TryGetUDTVariableAttributes(string udtVariableName, out IFieldEncapsulationAttributes rule)
        {
            rule = null;
            if (_udtVariableEncapsulationAttributes.ContainsKey(udtVariableName))
            {
                rule = _udtVariableEncapsulationAttributes[udtVariableName];
                return true;
            }
            return false;
        }

        private void AddUDTEncapsulationFields(IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> udtFieldToTypeMap)
        {
            foreach (var udtField in udtFieldToTypeMap.Keys)
            {
                var udtEncapsulation = DecorateUDTVariableDeclaration(udtField);
                AddEncapsulationField(udtEncapsulation);

                foreach (var udtMember in UdtFieldMembers(udtField))
                {
                    AddEncapsulationField(DecorateUDTMember(udtMember, udtEncapsulation as EncapsulateUserDefinedType));
                }
            }
        }

        public EncapsulateFieldDeclaration this[Declaration declaration]
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

        private EncapsulateFieldDeclaration this[string encapsulatedFieldIdentifier]
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

        public IEnumerable<EncapsulateFieldDeclaration> EncapsulateFieldDeclarations => _encapsulateFieldDeclarations.Values;

        public void AddEncapsulationField(EncapsulateFieldDeclaration target)
        {
            if (_encapsulateFieldDeclarations.ContainsKey(target.Declaration))
            {
                _encapsulateFieldDeclarations[target.Declaration] = target;
                return;
            }
            _encapsulateFieldDeclarations.Add(target.Declaration, target);
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

        private string BuildPropertiesTextBlock(Declaration target)
        {
            var attributes = this[target].EncapsulationAttributes;
            var generator = new PropertyGenerator
            {
                PropertyName = attributes.PropertyName,
                AsTypeName = attributes.AsTypeName,
                BackingField = attributes.NewFieldName,
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

        //private EncapsulateFieldDeclaration DefaultTarget => _userSelectedEncapsulationField;

        //Only used by tests....so far
        public void UpdateEncapsulationField(IUDTFieldEncapsulationAttributes attributes)
        {
            this[attributes.FieldName].EncapsulationAttributes = attributes;
            foreach ((string Name, bool Encapsulate) in attributes.MemberFlags)
            {
                this[Name].EncapsulationAttributes.IsFlaggedToEncapsulate = Encapsulate;
            }
        }

        //Only used by tests....so far
        public void UpdateEncapsulationField(string variableName, string propertyName, bool encapsulateFlag, string newFieldName = null)
        {
            var target = this[variableName];
            target.EncapsulationAttributes.FieldName = variableName;
            target.EncapsulationAttributes.PropertyName = propertyName;
            target.EncapsulationAttributes.NewFieldName = newFieldName ?? variableName;
            target.EncapsulationAttributes.IsFlaggedToEncapsulate = encapsulateFlag;
            this[variableName] = target;
        }

        //Only used by tests....so far
        public void UpdateEncapsulationField(IFieldEncapsulationAttributes attributes)
        {
            var target = this[attributes.FieldName];
            target.EncapsulationAttributes.FieldName = attributes.FieldName;
            target.EncapsulationAttributes.PropertyName = attributes.PropertyName;
            target.EncapsulationAttributes.NewFieldName = attributes.NewFieldName ?? attributes.FieldName;
            target.EncapsulationAttributes.IsFlaggedToEncapsulate = attributes.IsFlaggedToEncapsulate;
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

        public string ParameterName
        {
            get => _userSelectedEncapsulationField.EncapsulationAttributes.ParameterName;
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
            => _userSelectedEncapsulationField.EncapsulationAttributes.IsValueType || _userSelectedEncapsulationField.EncapsulationAttributes.IsVariantType;

        public bool CanImplementSet
            => !_userSelectedEncapsulationField.EncapsulationAttributes.IsValueType && _userSelectedEncapsulationField.EncapsulationAttributes.IsVariantType;
    }
}
