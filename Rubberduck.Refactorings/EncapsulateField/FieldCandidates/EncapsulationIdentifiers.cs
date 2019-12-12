using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.IO;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulationIdentifiers
    {
        private static string DEFAULT_WRITE_PARAMETER = "value";

        private KeyValuePair<string, string> _fieldAndProperty;
        private string _targetIdentifier;
        private string _defaultPropertyName;
        private string _setLetParameter;

        public EncapsulationIdentifiers(Declaration target)
            : this(target.IdentifierName) { }

        public EncapsulationIdentifiers(string field)
        {
            _targetIdentifier = field;
            _defaultPropertyName = field.Capitalize();
            DefaultNewFieldName = (field.UnCapitalize()).IncrementIdentifier();
            _fieldAndProperty = new KeyValuePair<string, string>(DefaultNewFieldName, _defaultPropertyName);
            _setLetParameter = DEFAULT_WRITE_PARAMETER;
        }

        public string TargetFieldName => _targetIdentifier;

        public string DefaultNewFieldName { private set; get; }

        public string Field
        {
            get => _fieldAndProperty.Key;
            set
            {
                _fieldAndProperty = new KeyValuePair<string, string>(value, _fieldAndProperty.Value);
                 SetNonConflictParameterName();
            }
        }

        public string Property
        {
            get => _fieldAndProperty.Value;
            set
            {
                var fieldIdentifier = Field.EqualsVBAIdentifier(value)
                        ? DefaultNewFieldName
                        : _targetIdentifier;

                _fieldAndProperty = new KeyValuePair<string, string>(_fieldAndProperty.Key, value);

                SetNonConflictParameterName();
            }
        }

        public string SetLetParameter => _setLetParameter;

        private void SetNonConflictParameterName()
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
