using Rubberduck.Parsing.Symbols;
using System;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IClientEditableFieldEncapsulationAttributes
    {
        string FieldName { get; }
        string NewFieldName { get; set; }
        string PropertyName { get; set; }
        bool ReadOnly { get; set; }
        bool EncapsulateFlag { get; set; }
    }

    public interface IFieldEncapsulationAttributes : IClientEditableFieldEncapsulationAttributes
    {
        string FieldReadWriteIdentifier { get; }
        string AsTypeName { get; set; }
        string ParameterName { get; set; }
        bool CanImplementLet { get; set; }
        bool CanImplementSet { get; set; }
        bool ImplementLetSetterType { get; set; }
        bool ImplementSetSetterType { get; set; }
        Func<string> FieldReadWriteIdentifierFunc { set; get; }
    }

    public struct ClientEncapsulationAttributes : IClientEditableFieldEncapsulationAttributes
    {
        public ClientEncapsulationAttributes(string targetName)
        {
            _fieldName = targetName;
            PropertyName = $"{char.ToUpperInvariant(targetName[0]) + targetName.Substring(1, targetName.Length - 1)}";
            NewFieldName = $"{targetName}1";
            ReadOnly = false;
            EncapsulateFlag = false;
        }

        private string _fieldName;
        public string FieldName => _fieldName;
        public string NewFieldName { get; set; }
        public string PropertyName { get; set; }
        public bool ReadOnly { get; set; }
        public bool EncapsulateFlag { get; set; }
    }

    public class FieldEncapsulationAttributes : IFieldEncapsulationAttributes
    {
        private static string DEFAULT_LET_PARAMETER = "value";

        public FieldEncapsulationAttributes(Declaration target, string newFieldName = null, string parameterName = null)
        {
            var defaults = new ClientEncapsulationAttributes(target.IdentifierName);
            FieldName = target.IdentifierName;
            NewFieldName = defaults.NewFieldName;
            PropertyName = defaults.PropertyName;
            AsTypeName = target.AsTypeName;
            ParameterName = parameterName ?? DEFAULT_LET_PARAMETER;
            FieldReadWriteIdentifierFunc = () => NewFieldName;
        }


        public string FieldName { get; private set; }

        private string _newFieldName;
        public string NewFieldName
        {
            get => EncapsulateFlag ? _newFieldName : FieldName;
            set => _newFieldName = value;
        }

        public string FieldReadWriteIdentifier => FieldReadWriteIdentifierFunc();
        public Func<string> FieldReadWriteIdentifierFunc { set; get; }
        public string PropertyName { get; set; }
        public string AsTypeName { get; set; }
        public string ParameterName { get; set; } = DEFAULT_LET_PARAMETER;
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
