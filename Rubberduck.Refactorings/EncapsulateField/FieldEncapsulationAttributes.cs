using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IClientEditableFieldEncapsulationAttributes
    {
        string FieldName { get; }
        string PropertyName { get; set; }
        bool ReadOnly { get; set; }
        bool EncapsulateFlag { get; set; }
    }

    public interface IFieldEncapsulationAttributes : IClientEditableFieldEncapsulationAttributes
    {
        string NewFieldName { get; } // set; }
        string FieldReadWriteIdentifier { get; }
        string AsTypeName { get; set; }
        string ParameterName { get;}
        bool CanImplementLet { get; set; } //TODO: Can these go away?
        bool CanImplementSet { get; set; } //TODO: Can these go away?
        bool ImplementLetSetterType { get; set; }
        bool ImplementSetSetterType { get; set; }
        Func<string> FieldReadWriteIdentifierFunc { set; get; }
    }

    public struct ClientEncapsulationAttributes : IClientEditableFieldEncapsulationAttributes
    {
        public ClientEncapsulationAttributes(string targetName)
        {
            _identifiers = new EncapsulationIdentifiers(targetName);
            //_fieldName = targetName;
            //PropertyName = _identifiers.Property; // $"{char.ToUpperInvariant(targetName[0]) + targetName.Substring(1, targetName.Length - 1)}";
            //NewFieldName = _identifiers.Field; // $"{char.ToLowerInvariant(targetName[0]) + targetName.Substring(1, targetName.Length - 1)}1";
            ReadOnly = false;
            EncapsulateFlag = false;
        }

        //private string _fieldName;
        private EncapsulationIdentifiers _identifiers;
        public string FieldName => _identifiers.TargetFieldName;
        public string NewFieldName
        {
            get => _identifiers.Field;
            set => _identifiers.Field = value;
        }
        public string PropertyName //{ get; set; }
        {
            get => _identifiers.Property;
            set => _identifiers.Property = value;
        }
        public bool ReadOnly { get; set; }
        public bool EncapsulateFlag { get; set; }
    }

    public class EncapsulationIdentifiers
    {
        private static string DEFAULT_WRITE_PARAMETER = "value";

        private KeyValuePair<string, string> _fieldAndProperty;
        private string _targetIdentifier;
        //private Declaration _target;
        private string _defaultPropertyName;
        private string _setLetParameter;

        public EncapsulationIdentifiers(Declaration target)
            : this(target.IdentifierName) { }

        public EncapsulationIdentifiers(string field)
        {
            string Capitalize(string input) => $"{char.ToUpperInvariant(input[0]) + input.Substring(1, input.Length - 1)}";
            string UnCapitalize(string input) => $"{char.ToLowerInvariant(input[0]) + input.Substring(1, input.Length - 1)}";

            _targetIdentifier = field;
            _defaultPropertyName = Capitalize(field);
            _fieldAndProperty = new KeyValuePair<string, string>($"{UnCapitalize(field)}1", _defaultPropertyName);
            _setLetParameter = DEFAULT_WRITE_PARAMETER;
        }

        public string TargetFieldName => _targetIdentifier;

        public string Field
        {
            get => _fieldAndProperty.Key;
            set
            {
                _fieldAndProperty = new KeyValuePair<string, string>(value, _fieldAndProperty.Value);
                GenerateNonConflictParamIdentifier();
            }
        }

        public string Property
        {
            get => _fieldAndProperty.Value;
            set
            {
                //Reverts to original field name if user modifies generated Property name
                if (!IsVBAEquivalentName(value, Field))
                {
                    _fieldAndProperty = new KeyValuePair<string, string>(_targetIdentifier, value);
                }
                GenerateNonConflictParamIdentifier();
            }
        }

        public string SetLetParameter => _setLetParameter;

        private bool IsVBAEquivalentName(string input, string existingName)
            => input.Equals(existingName, StringComparison.InvariantCultureIgnoreCase);

        private void GenerateNonConflictParamIdentifier()
        {
            if (IsVBAEquivalentName(Field, SetLetParameter))
            {
                _setLetParameter = $"{Field}_{DEFAULT_WRITE_PARAMETER}";
            }

            if (IsVBAEquivalentName(Property, SetLetParameter))
            {
                _setLetParameter = $"{Property}_{Field}_{DEFAULT_WRITE_PARAMETER}";
            }
        }
    }

    public class FieldEncapsulationAttributes : IFieldEncapsulationAttributes
    {

        public FieldEncapsulationAttributes(Declaration target, string newFieldName = null)
        {
            var defaults = new ClientEncapsulationAttributes(target.IdentifierName);
            FieldName = target.IdentifierName;
            //NewFieldName = defaults.NewFieldName;
            //PropertyName = defaults.PropertyName;
            AsTypeName = target.AsTypeName;
            //ParameterName = parameterName ?? SetLetParameter;
            FieldReadWriteIdentifierFunc = () => NewFieldName;
            _fieldAndProperty = new EncapsulationIdentifiers(target);
        }

        private EncapsulationIdentifiers _fieldAndProperty;

        public string FieldName { private set; get; }

        //private string _newFieldName;
        public string NewFieldName
        {
            get => _fieldAndProperty.Field;
            set => _fieldAndProperty.Field = value;
        }

        //public string PropertyName { get; set; }
        public string PropertyName
        {
            get => _fieldAndProperty.Property;
            set => _fieldAndProperty.Property = value;
        }

        public string FieldReadWriteIdentifier => FieldReadWriteIdentifierFunc();

        public Func<string> FieldReadWriteIdentifierFunc { set; get; }
        public string AsTypeName { get; set; }
        public string ParameterName => _fieldAndProperty.SetLetParameter;
        public bool ReadOnly { get; set; }

        private bool _implLet;
        public bool ImplementLetSetterType { get => !ReadOnly && _implLet; set => _implLet = value; }

        private bool _implSet;
        public bool ImplementSetSetterType { get => !ReadOnly && _implSet; set => _implSet = value; }

        public bool EncapsulateFlag { get; set; }
        public bool CanImplementLet { get; set; }
        public bool CanImplementSet { get; set; }
    }
}
