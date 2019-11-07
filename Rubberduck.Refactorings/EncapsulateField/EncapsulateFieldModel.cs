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
    public interface IEncapsulateFieldAttributes
    {
        string FieldName { set; get; }
        string PropertyName { set; get; }
        string ParameterName { set; get; }
        bool ImplementLetSetterType { set; get; }
        bool ImplementSetSetterType { set; get; }
        bool CanImplementLet { get; }
        bool CanImplementSet { get; }
    }

    public struct EncapsulateUDTVariableRule
    {
        public EncapsulateUDTVariableRule(string variableName, string propertyName = null, string asTypeName = "Variant")
        {
            Attributes = new EncapsulationAttributes(variableName, propertyName ?? $"{variableName}_{Tokens.Type}", asTypeName);
            EncapsulateVariable = true;
            EncapsulateAllUDTMembers = false;
        }
        public EncapsulationAttributes Attributes { set; get; }
        public string VariableName => Attributes?.FieldName ?? string.Empty;
        public bool EncapsulateVariable { set; get; }
        public bool EncapsulateAllUDTMembers { set; get; }
    }

    public class EncapsulationAttributes : IEncapsulateFieldAttributes
    {
        private static string DEFAULT_LET_PARAMETER = "value";
        public EncapsulationAttributes()
        {
            FieldName = string.Empty;
            PropertyName = string.Empty;
            AsTypeName = string.Empty;
            ParameterName = DEFAULT_LET_PARAMETER;
            ImplementLetSetterType = false;
            ImplementSetSetterType = false;
            CanImplementLet = false;
            CanImplementSet = false;
            Encapsulate = false;
        }

        public EncapsulationAttributes(string fieldName, string propertyName, string asTypeName)
        {
            FieldName = fieldName;
            PropertyName = propertyName;
            AsTypeName = asTypeName;
            ParameterName = DEFAULT_LET_PARAMETER;
            ImplementLetSetterType = false;
            ImplementSetSetterType = false;
            CanImplementLet = true;
            CanImplementSet = true;
            Encapsulate = false;
        }

        public EncapsulationAttributes(Declaration target)
        {
            FieldName = target.IdentifierName;
            PropertyName = target.IdentifierName;
            AsTypeName = target.AsTypeName;
            ParameterName = DEFAULT_LET_PARAMETER;
            ImplementLetSetterType = false;
            ImplementSetSetterType = false;
            Encapsulate = false;
            IsVariant = target.AsTypeName?.Equals(Tokens.Variant) ?? true;
            IsValueType = !IsVariant && (SymbolList.ValueTypes.Contains(target.AsTypeName) ||
                                             target.DeclarationType == DeclarationType.Enumeration);
            CanImplementLet = IsValueType || IsVariant;
            CanImplementSet = !(IsValueType || IsVariant);
        }

        public EncapsulationAttributes(EncapsulationAttributes attributes)
        {
            FieldName = attributes.FieldName;
            PropertyName = attributes.PropertyName;
            AsTypeName = attributes.AsTypeName;
            ParameterName = attributes.ParameterName;
            ImplementLetSetterType = attributes.ImplementLetSetterType;
            ImplementSetSetterType = attributes.ImplementSetSetterType;
            CanImplementLet = attributes.CanImplementLet;
            CanImplementSet = attributes.CanImplementSet;
            IsVariant = attributes.IsVariant;
            IsValueType = attributes.IsValueType;
            Encapsulate = false;
        }

        public string FieldName { get; set; }
        public string PropertyName { get; set; }
        public string AsTypeName { get; set; }
        public string ParameterName { get; set; }
        public bool ImplementLetSetterType { get; set; }
        public bool ImplementSetSetterType { get; set; }

        public bool CanImplementLet { get; set; }
        public bool CanImplementSet { get; set; }
        public bool Encapsulate { get; set; }
        public bool IsValueType { private set;  get; }
        public bool IsVariant { private set;  get; }
    }

    public class EncapsulateFieldModel : IRefactoringModel
    {
        private Dictionary<Declaration,(Declaration, IEnumerable<Declaration>)> _udts = new Dictionary<Declaration, (Declaration, IEnumerable<Declaration>)>();
        private readonly IIndenter _indenter;
        private IList<EncapsulateFieldDeclaration> _fields = new List<EncapsulateFieldDeclaration>();
        private IList<Declaration> _nonUdtVariables = new List<Declaration>();
        private Dictionary<string,Declaration> _nonUdtVariablesByName = new Dictionary<string, Declaration>();
        private EncapsulateFieldDeclaration _selectedTarget;

        public EncapsulateFieldModel(Declaration target, IEnumerable<Declaration> nonUdtVariables, IEnumerable<(Declaration udtVariable, Declaration udt, IEnumerable<Declaration> udtMembers)> udtTuples, IIndenter indenter)
        {
            _indenter = indenter;

            foreach (var variable in _nonUdtVariables)
            {
                LoadNonUDTVariable(variable);
            }

            foreach (var udtTuple in udtTuples)
            {
                LoadUDTVariable(udtTuple.udtVariable, udtTuple.udt, udtTuple.udtMembers);
            }

            var selectedField = new EncapsulateFieldDeclaration(target);

            if (UDTVariables.Contains(target))
            {
                _selectedTarget = new UserDefinedTypeField(selectedField);
            }
            else if (selectedField.EncapsulationAttributes.IsValueType)
            {
                _selectedTarget = new EncapsulatedValueType(selectedField);
            }
            else if (selectedField.EncapsulationAttributes.IsVariant)
            {
                _selectedTarget = new EncapsulatedVariantType(selectedField);
            }

            _selectedTarget.EncapsulationAttributes.Encapsulate = true;

            AddEncapsulationTarget(_selectedTarget);
        }

        private Dictionary<Declaration, EncapsulationAttributes> _encapsulationAttributesByFieldDeclaration = new Dictionary<Declaration, EncapsulationAttributes>();
        private Dictionary<string, EncapsulationAttributes> _encapsulationAttributesByFieldIdentifier = new Dictionary<string, EncapsulationAttributes>();

        private Dictionary<string, EncapsulateUDTVariableRule> _udtVariableRules = new Dictionary<string, EncapsulateUDTVariableRule>();

        public void AddUDTVariableRule(EncapsulateUDTVariableRule udtRule)
        {
            if (_udtVariableRules.TryGetValue(udtRule.VariableName, out _))
            {
                _udtVariableRules[udtRule.VariableName] = udtRule;
                return;
            }
            _udtVariableRules.Add(udtRule.VariableName, udtRule);
        }

        public bool TryGetUDTVariableRule(string udtVariableName, out EncapsulateUDTVariableRule rule)
        {
            rule = new EncapsulateUDTVariableRule();
            if (_udtVariableRules.ContainsKey(udtVariableName))
            {
                rule = _udtVariableRules[udtVariableName];
                return true;
            }
            return false;
        }

        public IEnumerable<EncapsulateUDTVariableRule> UDTVariableRules => _udtVariableRules.Values;

        public IEnumerable<Declaration> UDTVariables => _udts.Keys;

        private void LoadUDTVariable(Declaration udtVariable, Declaration udt, IEnumerable<Declaration> udtMembers)
        {
            _udts.Add(udtVariable, (udt, udtMembers));
            var udtVariableRule = new EncapsulateUDTVariableRule(udtVariable.IdentifierName);
            AddUDTVariableRule(udtVariableRule);
        }

        private void LoadNonUDTVariable(Declaration variable)
        {
            _nonUdtVariables.Add(variable);
            _nonUdtVariablesByName.Add(variable.IdentifierName, variable);
        }

        public IEnumerable<Declaration> GetUdtMembers(Declaration udtVariable)
        {
            if (_udts.TryGetValue(udtVariable, out var value))
            {
                return value.Item2;
            }
            return Enumerable.Empty<Declaration>();
        }

        public IEncapsulateFieldAttributes this[string fieldIdentifier] => _encapsulationAttributesByFieldIdentifier[fieldIdentifier];

        public IEncapsulateFieldAttributes this[Declaration field] => _encapsulationAttributesByFieldDeclaration[field];

        public IEnumerable<Declaration> EncapsulationTargets => _fields.Select(f => f.Declaration); // _encapsulationAttributesByFieldDeclaration.Keys;

        public void AddEncapsulationTarget(EncapsulateFieldDeclaration target) => _fields.Add(target);

        public IList<string> PropertiesContent
        {
            get
            {
                var textBlocks = new List<string>();
                foreach (var field in _fields)
                {
                    switch (field)
                    {
                        case EncapsulatedUserDefinedMemberValueType udtMemberType:
                            if (udtMemberType.EncapsulationAttributes.Encapsulate)
                            {
                                textBlocks.Add(BuildPropertiesTextBlock(udtMemberType.Declaration));
                            }
                            break;
                        case UserDefinedTypeField udtType:
                            if (udtType.EncapsulationAttributes.Encapsulate)
                            {
                                textBlocks.Add(BuildPropertiesTextBlock(udtType.Declaration));
                            }
                            break;
                        default:
                            textBlocks.Add(BuildPropertiesTextBlock(field.Declaration));
                            break;
                    }
                }
                return textBlocks;
            }
        }

        private string BuildPropertiesTextBlock(Declaration target)
        {
            var attributes = EncapsulationAttributes(target);
            var generator = new PropertyGenerator
            {
                PropertyName = attributes.PropertyName,
                AsTypeName = attributes.AsTypeName,
                BackingField = attributes.FieldName,
                ParameterName = attributes.ParameterName,
                GenerateSetter = attributes.ImplementSetSetterType,
                GenerateLetter = attributes.ImplementLetSetterType
            };

            return GetPropertyText(generator);
        }

        private string BuildUDTMemberPropertiesTextBlock(Declaration udtMember)
        {
            var attributes = EncapsulationAttributes(udtMember);

            var udtVariable = _udts.Keys.Where(k => _udts[k].Item2.Contains(udtMember)).SingleOrDefault();

            var generator = new PropertyGenerator
            {
                PropertyName = attributes.PropertyName,
                AsTypeName = attributes.AsTypeName,
                BackingField = $"{udtVariable.IdentifierName}.{udtMember.IdentifierName}",
                ParameterName = attributes.ParameterName,
                GenerateSetter = udtMember.IsObject, //model.ImplementSetSetterType,
                GenerateLetter = !udtMember.IsObject //model.ImplementLetSetterType
            };

            return GetPropertyText(generator);
        }

        private string GetPropertyText(PropertyGenerator generator)
        {
            var propertyTextLines = generator.AllPropertyCode.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            return string.Join(Environment.NewLine, _indenter.Indent(propertyTextLines, true));
        }

        public EncapsulationAttributes EncapsulationAttributes(Declaration target)
        {
            foreach ( var field in _fields)
            {
                if (field.Declaration == target)
                {
                    return field.EncapsulationAttributes;
                }
            }
            return null;
        }

        private EncapsulateFieldDeclaration DefaultTarget => _selectedTarget;
        
        public Declaration TargetDeclaration
        {
            get => DefaultTarget.Declaration;
            set => _selectedTarget = new EncapsulateFieldDeclaration(value);
        }

        public string PropertyName
        {
            get => DefaultTarget.EncapsulationAttributes.PropertyName;
            set => DefaultTarget.EncapsulationAttributes.PropertyName = value;
        }

        public string ParameterName
        {
            get => DefaultTarget.EncapsulationAttributes.ParameterName;
            set => DefaultTarget.EncapsulationAttributes.ParameterName = value;
        }

        public bool ImplementLetSetterType
        {
            get => DefaultTarget.EncapsulationAttributes.ImplementLetSetterType;
            set => DefaultTarget.EncapsulationAttributes.ImplementLetSetterType = value;
        }

        public bool ImplementSetSetterType
        {
            get => DefaultTarget.EncapsulationAttributes.ImplementSetSetterType;
            set => DefaultTarget.EncapsulationAttributes.ImplementSetSetterType = value;
        }

        public bool CanImplementLet
            => DefaultTarget.EncapsulationAttributes.CanImplementLet;

        public bool CanImplementSet
            => DefaultTarget.EncapsulationAttributes.CanImplementSet;
    }
}
