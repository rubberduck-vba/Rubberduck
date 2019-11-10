using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IUserModifiableFieldEncapsulationAttributes
    {
        string FieldName { get; set; }
        string NewFieldName { get; set; }
        string PropertyName { get; set; }
        bool ImplementLetSetterType { get; set; }
        bool ImplementSetSetterType { get; set; }
        bool EncapsulateFlag { get; set; }
    }

    public interface IFieldEncapsulationAttributes : IUserModifiableFieldEncapsulationAttributes
    {
        string AsTypeName { get; set; }
        string ParameterName { get; set; }
        bool CanImplementLet { get; set; }
        bool CanImplementSet { get; set; }
    }

    public class FieldEncapsulationAttributes : IFieldEncapsulationAttributes
    {
        private static string DEFAULT_LET_PARAMETER = "value";

        public FieldEncapsulationAttributes() { }

        public FieldEncapsulationAttributes(Declaration target, string newFieldName = null, string parameterName = null)
        {
            FieldName = target.IdentifierName;
            NewFieldName = newFieldName ?? target.IdentifierName;
            PropertyName = target.IdentifierName;
            AsTypeName = target.AsTypeName;
            ParameterName = parameterName ?? DEFAULT_LET_PARAMETER;
        }

        public string FieldName { get; set; }

        private string _newFieldName;
        public string NewFieldName
        {
            get => _newFieldName ?? FieldName;
            set => _newFieldName = value;
        }
        public string PropertyName { get; set; }
        public string AsTypeName { get; set; }
        public string ParameterName { get; set; } = DEFAULT_LET_PARAMETER;

        public bool ImplementLetSetterType { get; set; }
        public bool ImplementSetSetterType { get; set; }
        public bool EncapsulateFlag { get; set; }
        public bool CanImplementLet { get; set; }
        public bool CanImplementSet { get; set; }
    }

    public class UDTFieldEncapsulationAttributes : IFieldEncapsulationAttributes
    {
        private IFieldEncapsulationAttributes _attributes { get; }
        private Dictionary<string, bool> _memberEncapsulationFlags = new Dictionary<string, bool>();

        public UDTFieldEncapsulationAttributes(FieldEncapsulationAttributes attributes) //, IEnumerable<string> udtMemberNames)
        {
            _attributes = attributes;
        }

        public string FieldName
        {
            get => _attributes.FieldName;
            set => _attributes.FieldName = value;
        }

        public string PropertyName
        {
            get => _attributes.PropertyName;
            set => _attributes.PropertyName = value;
        }

        public string NewFieldName
        {
            get => _attributes.NewFieldName;
            set => _attributes.NewFieldName = value;
        }

        public string AsTypeName
        {
            get => _attributes.AsTypeName;
            set => _attributes.AsTypeName = value;
        }
        public string ParameterName
        {
            get => _attributes.ParameterName;
            set => _attributes.ParameterName = value;
        }
        public bool ImplementLetSetterType
        {
            get => _attributes.ImplementLetSetterType;
            set => _attributes.ImplementLetSetterType = value;
        }

        public bool ImplementSetSetterType
        {
            get => _attributes.ImplementSetSetterType;
            set => _attributes.ImplementSetSetterType = value;
        }

        public bool EncapsulateFlag
        {
            get => _attributes.EncapsulateFlag;
            set => _attributes.EncapsulateFlag = value;
        }

        public bool CanImplementLet { get; set; }
        public bool CanImplementSet { get; set; }
    }
}
