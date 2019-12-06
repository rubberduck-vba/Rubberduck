using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.IO;

namespace Rubberduck.Refactorings.EncapsulateField
{
    //public struct EncapsulationAttributeIdentifier
    //{
    //    public EncapsulationAttributeIdentifier(string name, bool isImmutable = false)
    //    {
    //        Name = name;
    //        IsImmutable = isImmutable;
    //    }
    //    public string Name;
    //    public bool IsImmutable;
    //}

    public class EncapsulationIdentifiers
    {
        private static string DEFAULT_WRITE_PARAMETER = "value";

        private KeyValuePair<string, string> _fieldAndProperty;
        //private EncapsulationAttributeIdentifier _targetIdentifier;
        private string _targetIdentifier;
        private string _defaultPropertyName;
        private string _defaultFieldName;
        private string _setLetParameter;
        private bool _cannotEncapsulate;

        public EncapsulationIdentifiers(Declaration target)
            : this(target.IdentifierName) { }

        public EncapsulationIdentifiers(string field, bool cannotBeEncapsulated = false)
        {
            //_targetIdentifier = new EncapsulationAttributeIdentifier(field, true);
            _targetIdentifier = field;
            _defaultPropertyName = cannotBeEncapsulated ? $"{field}99" : field.Capitalize();
            _defaultFieldName = cannotBeEncapsulated ? field : $"{field.UnCapitalize()}1";
            _fieldAndProperty = new KeyValuePair<string, string>(_defaultFieldName, _defaultPropertyName);
            _setLetParameter = DEFAULT_WRITE_PARAMETER;
            _cannotEncapsulate = cannotBeEncapsulated;
        }

        public EncapsulationIdentifiers(string field, string fieldName, string propertyName)
        {
            //_targetIdentifier = new EncapsulationAttributeIdentifier(field);
            _targetIdentifier = field;
            _defaultPropertyName = field.Capitalize();
            _defaultFieldName = $"{field.UnCapitalize()}1";
            _fieldAndProperty = new KeyValuePair<string, string>(fieldName, propertyName);
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
                var fieldIdentifier = Field.EqualsVBAIdentifier(value)
                        ? _defaultFieldName
                        : _targetIdentifier;

                _fieldAndProperty = new KeyValuePair<string, string>(fieldIdentifier, value);

                GenerateNonConflictParamIdentifier();
            }
        }

        public string SetLetParameter => _setLetParameter;

        private void GenerateNonConflictParamIdentifier()
        {
            _setLetParameter = DEFAULT_WRITE_PARAMETER;

            if (!(Field.EqualsVBAIdentifier(DEFAULT_WRITE_PARAMETER)
                    || Property.EqualsVBAIdentifier(DEFAULT_WRITE_PARAMETER)))
            {
                return;
            }

            if (Field.EqualsVBAIdentifier(SetLetParameter))
            {
                _setLetParameter = $"{Field}_{DEFAULT_WRITE_PARAMETER}";
            }

            if (Property.EqualsVBAIdentifier(SetLetParameter))
            {
                _setLetParameter = $"{Property}_{Field}_{DEFAULT_WRITE_PARAMETER}";
            }
        }
    }
}
