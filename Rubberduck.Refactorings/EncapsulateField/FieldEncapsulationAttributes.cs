using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IFieldEncapsulationAttributes
    {
        string Identifier { get; }
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
        QualifiedModuleName QualifiedModuleName { get; }

        //IEncapsulatedField
        //Declaration Declaration { get; }
        //DeclarationType DeclarationType { get; }
        //dup        //string TargetID { get; }
        //IFieldEncapsulationAttributes EncapsulationAttributes { set; get; }
        //dup        //bool IsReadOnly { set; get; }
        //bool CanBeReadWrite { set; get; }
        //dup        //string PropertyName { set; get; }
        //dup        //bool EncapsulateFlag { set; get; }
        //dup        //string NewFieldName { get; }
        //dup        //string AsTypeName { get; }
        //bool IsUDTMember { set; get; }
        //bool HasValidEncapsulationAttributes { get; }
        //dup//QualifiedModuleName QualifiedModuleName { get; }
        //IEnumerable<IdentifierReference> References { get; }

    }

    public class UnselectableField : IFieldEncapsulationAttributes
    {
        private const string neverUsed = "x_x_x_x_x_x_x";
        private IEncapsulateFieldNamesValidator _validator;
        private QualifiedModuleName _qmn;

        public UnselectableField(string identifier, string asTypeName, QualifiedModuleName qmn, IEncapsulateFieldNamesValidator validator)
        {
            //_fieldAndProperty = new EncapsulationIdentifiers(identifier);
            _qmn = qmn;
            _validator = validator;
            Identifier = identifier;
            NewFieldName = identifier;
            AsTypeName = asTypeName;
            FieldReferenceExpressionFunc = () => NewFieldName;
        }

        public IFieldEncapsulationAttributes ApplyNewFieldName(string newFieldName)
        {
            NewFieldName = newFieldName;
            return this;
        }

        public string Identifier { private set; get; }

        public string NewFieldName { set; get; }

        string _tossString;
        public string PropertyName { set => _tossString = value; get => $"{neverUsed}{Identifier}_{neverUsed}"; }

        public string FieldReferenceExpression => FieldReferenceExpressionFunc();

        public Func<string> FieldReferenceExpressionFunc { set; get; }

        public string AsTypeName { get; set; }
        public string ParameterName => neverUsed;

        private bool _toss;
        public bool ReadOnly { get; set; } = false;

        public bool EncapsulateFlag { get => false; set => _toss = value; }
        public bool ImplementLetSetterType { get => false; set => _toss = value; }
        public bool ImplementSetSetterType { get => false; set => _toss = value; }

        public bool FieldNameIsExemptFromValidation => false; // NewFieldName.EqualsVBAIdentifier(Identifier);
        public QualifiedModuleName QualifiedModuleName => _qmn;
    }

    public class FieldEncapsulationAttributes : IFieldEncapsulationAttributes
    {
        private QualifiedModuleName _qmn;
        public FieldEncapsulationAttributes(Declaration target)
        {
            _fieldAndProperty = new EncapsulationIdentifiers(target);
            Identifier = target.IdentifierName;
            AsTypeName = target.AsTypeName;
            _qmn = target.QualifiedModuleName;
            FieldReferenceExpressionFunc = () => NewFieldName;
        }

        public FieldEncapsulationAttributes(string identifier, string asTypeName)
        {
            _fieldAndProperty = new EncapsulationIdentifiers(identifier);
            AsTypeName = asTypeName;
            FieldReferenceExpressionFunc = () => NewFieldName;
        }

        private EncapsulationIdentifiers _fieldAndProperty;

        public string Identifier { private set; get; }

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

        public bool FieldNameIsExemptFromValidation => NewFieldName.EqualsVBAIdentifier(Identifier);
        public QualifiedModuleName QualifiedModuleName => _qmn;
    }
}
