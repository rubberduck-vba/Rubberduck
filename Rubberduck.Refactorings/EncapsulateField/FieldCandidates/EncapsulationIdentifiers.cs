using Rubberduck.Common;
using System.Collections.Generic;
using Rubberduck.Refactorings.EncapsulateField.Extensions;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulationIdentifiers
    {
        private KeyValuePair<string, string> _fieldAndProperty;
        private string _targetIdentifier;

        public EncapsulationIdentifiers(string field, IValidateVBAIdentifiers identifierValidator)
        {
            _targetIdentifier = field;

            DefaultPropertyName = field.CapitalizeFirstLetter();
            DefaultNewFieldName = (field.ToLowerCaseFirstLetter()).IncrementEncapsulationIdentifier();

            if (field.TryMatchHungarianNotationCriteria(out var nonHungarianName))
            {
                if (identifierValidator.IsValidVBAIdentifier(nonHungarianName, out _))
                {
                    DefaultPropertyName = nonHungarianName;
                    DefaultNewFieldName = field;
                }
            }
            else if (field.StartsWith("m_"))
            {
                var propertyName = field.Substring(2).CapitalizeFirstLetter();
                if (identifierValidator.IsValidVBAIdentifier(propertyName, out _))
                {
                    DefaultPropertyName = propertyName;
                    DefaultNewFieldName = field;
                }
            }

            _fieldAndProperty = new KeyValuePair<string, string>(DefaultNewFieldName, DefaultPropertyName);
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
            }
        }

        public string Property
        {
            get => _fieldAndProperty.Value;
            set
            {
                _fieldAndProperty = new KeyValuePair<string, string>(_fieldAndProperty.Key, value);
            }
        }
    }
}
