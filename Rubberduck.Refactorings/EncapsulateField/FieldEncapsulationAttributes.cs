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
        //string SetGet_LHSField { get; set; }
        string AsTypeName { get; set; }
        string ParameterName { get; set; }
        bool CanImplementLet { get; set; }
        bool CanImplementSet { get; set; }
    }

    public class FieldEncapsulationAttributes : IFieldEncapsulationAttributes
    {
        private static string DEFAULT_LET_PARAMETER = "value";

        //private string _defaultBackingFieldName;
        //private string _defaultPropertyName;
        //private string _defaultParameterName = DEFAULT_LET_PARAMETER;

        public FieldEncapsulationAttributes() { }

        //public FieldEncapsulationAttributes(Declaration target)
        //{
        //    FieldName = target.IdentifierName;
        //    (string newFieldName, string propertyName) = GenerateDefaultNames(target.IdentifierName);
        //    SetGet_LHSField = newFieldName;
        //    NewFieldName = newFieldName;
        //    PropertyName = propertyName;
        //    AsTypeName = target.AsTypeName;
        //    ParameterName = DEFAULT_LET_PARAMETER;
        //}

        public FieldEncapsulationAttributes(Declaration target, string newFieldName = null, string parameterName = null)
        {
            FieldName = target.IdentifierName;
            NewFieldName = newFieldName ?? target.IdentifierName;
            PropertyName = target.IdentifierName;
            AsTypeName = target.AsTypeName;
            ParameterName = parameterName ?? DEFAULT_LET_PARAMETER;
        }

        public (string defaultField, string defaultProperty) GenerateDefaultNames(string targetName)
        {
            var defaultBackingFieldName = targetName + "1";
            var defaultPropertyName = string.Empty;
            if (char.IsLower(targetName, 0))
            {
                defaultPropertyName = $"{char.ToUpperInvariant(targetName[0]) + targetName.Substring(1, targetName.Length - 1)}";
            }
            else
            {
                defaultPropertyName = targetName;
            }
            return (defaultBackingFieldName, defaultPropertyName);
        }

        //private bool IsCamelCase(string value) => char.IsLower(value, 0) && char.IsUpper(value, 1);

        public string FieldName { get; set; }

        private string _newFieldName;
        public string NewFieldName
        {
            get => _newFieldName ?? FieldName;
            set => _newFieldName = value;
        }

        //TODO: Set BackingField to udtVariable.udtMember somehow.  Otherwise, BackingField == NewField
        //public string SetGet_LHSField { get; set; }
        public string PropertyName { get; set; }
        public string AsTypeName { get; set; }
        public string ParameterName { get; set; } = DEFAULT_LET_PARAMETER;

        public bool ImplementLetSetterType { get; set; }
        public bool ImplementSetSetterType { get; set; }
        public bool EncapsulateFlag { get; set; }
        public bool CanImplementLet { get; set; }
        public bool CanImplementSet { get; set; }
    }

    //public class UDTFieldEncapsulationAttributesX : IFieldEncapsulationAttributes
    //{
    //    private IFieldEncapsulationAttributes _attributes { get; }
    //    private Dictionary<string, bool> _memberEncapsulationFlags = new Dictionary<string, bool>();

    //    public UDTFieldEncapsulationAttributesX(FieldEncapsulationAttributes attributes) //, IEnumerable<string> udtMemberNames)
    //    {
    //        _attributes = attributes;
    //    }

    //    public string FieldName
    //    {
    //        get => _attributes.FieldName;
    //        set => _attributes.FieldName = value;
    //    }

    //    public string PropertyName
    //    {
    //        get => _attributes.PropertyName;
    //        set => _attributes.PropertyName = value;
    //    }

    //    public string NewFieldName
    //    {
    //        get => _attributes.NewFieldName;
    //        set => _attributes.NewFieldName = value;
    //    }

    //    public string AsTypeName
    //    {
    //        get => _attributes.AsTypeName;
    //        set => _attributes.AsTypeName = value;
    //    }
    //    public string ParameterName
    //    {
    //        get => _attributes.ParameterName;
    //        set => _attributes.ParameterName = value;
    //    }
    //    public bool ImplementLetSetterType
    //    {
    //        get => _attributes.ImplementLetSetterType;
    //        set => _attributes.ImplementLetSetterType = value;
    //    }

    //    public bool ImplementSetSetterType
    //    {
    //        get => _attributes.ImplementSetSetterType;
    //        set => _attributes.ImplementSetSetterType = value;
    //    }

    //    public bool EncapsulateFlag
    //    {
    //        get => _attributes.EncapsulateFlag;
    //        set => _attributes.EncapsulateFlag = value;
    //    }

    //    public bool CanImplementLet { get; set; }
    //    public bool CanImplementSet { get; set; }
    //}
}
