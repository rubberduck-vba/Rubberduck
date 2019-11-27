using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IFieldEncapsulationAttributes
    {
        string TargetName { get; }
        string PropertyName { get; set; }
        bool ReadOnly { get; set; }
        bool EncapsulateFlag { get; set; }
        string NewFieldName { set;  get; }
        string FieldReferenceExpression { get; }
        string AsTypeName { get; set; }
        string ParameterName { get;}
        bool ImplementLetSetterType { get; set; }
        bool ImplementSetSetterType { get; set; }
        bool FieldNameIsExemptFromValidation { get; }
    }

    public class FieldEncapsulationAttributes : IFieldEncapsulationAttributes
    {

        public FieldEncapsulationAttributes(Declaration target)
        {
            _fieldAndProperty = new EncapsulationIdentifiers(target);
            TargetName = target.IdentifierName;
            AsTypeName = target.AsTypeName;
            FieldReferenceExpressionFunc = () => NewFieldName;
        }

        private EncapsulationIdentifiers _fieldAndProperty;

        public string TargetName { private set; get; }

        public string NewFieldName
        {
            get => _fieldAndProperty.Field;
            set => _fieldAndProperty.Field = value;
        }

        public string PropertyName
        {
            get => _fieldAndProperty.Property;
            set => _fieldAndProperty.Property = value;
        }

        public string FieldReferenceExpression => FieldReferenceExpressionFunc();

        public Func<string> FieldReferenceExpressionFunc { set; get; }

        public string AsTypeName { get; set; }
        public string ParameterName => _fieldAndProperty.SetLetParameter;
        public bool ReadOnly { get; set; }
        public bool EncapsulateFlag { get; set; }

        private bool _implLet;
        public bool ImplementLetSetterType { get => !ReadOnly && _implLet; set => _implLet = value; }

        private bool _implSet;
        public bool ImplementSetSetterType { get => !ReadOnly && _implSet; set => _implSet = value; }

        public bool FieldNameIsExemptFromValidation => NewFieldName.EqualsVBAIdentifier(TargetName);
    }
}
