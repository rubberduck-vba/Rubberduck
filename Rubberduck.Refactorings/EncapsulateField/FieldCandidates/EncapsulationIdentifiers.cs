using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.Extensions;
using System.Collections.Generic;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulationIdentifiers
    {
        private static string DEFAULT_WRITE_PARAMETER = EncapsulateFieldResources.DefaultPropertyParameter;

        private KeyValuePair<string, string> _fieldAndProperty;
        private string _targetIdentifier;
        private string _setLetParameter;

        public EncapsulationIdentifiers(Declaration target)
            : this(target.IdentifierName) { }

        public EncapsulationIdentifiers(string field)
        {
            _targetIdentifier = field;
            if (field.IsHungarianIdentifier(out var nonHungarianName))
            {
                DefaultPropertyName = nonHungarianName;
                DefaultNewFieldName = field;
            }
            else if (field.StartsWith("m_"))
            {
                DefaultPropertyName = field.Substring(2).Capitalize();
                DefaultNewFieldName = field;
            }
            else
            {
                DefaultPropertyName = field.Capitalize();
                DefaultNewFieldName = (field.UnCapitalize()).IncrementEncapsulationIdentifier();
            }
            _fieldAndProperty = new KeyValuePair<string, string>(DefaultNewFieldName, DefaultPropertyName);
            _setLetParameter = DEFAULT_WRITE_PARAMETER;
        }

        public string TargetFieldName => _targetIdentifier;

        public string DefaultPropertyName { private set; get; }

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
                _fieldAndProperty = new KeyValuePair<string, string>(_fieldAndProperty.Key, value);

                SetNonConflictParameterName();
            }
        }

        public string SetLetParameter => _setLetParameter;

        private void SetNonConflictParameterName()
        {
            _setLetParameter = DEFAULT_WRITE_PARAMETER;

            if (!(Field.IsEquivalentVBAIdentifierTo(DEFAULT_WRITE_PARAMETER)
                    || Property.IsEquivalentVBAIdentifierTo(DEFAULT_WRITE_PARAMETER)))
            {
                return;
            }

            if (Field.IsEquivalentVBAIdentifierTo(SetLetParameter))
            {
                _setLetParameter = $"{Field}_{DEFAULT_WRITE_PARAMETER}";
            }

            if (Property.IsEquivalentVBAIdentifierTo(SetLetParameter))
            {
                _setLetParameter = $"{Property}_{Field}_{DEFAULT_WRITE_PARAMETER}";
            }
        }
    }
}
