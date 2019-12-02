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
        bool IsReadOnly { get; set; }
        bool EncapsulateFlag { get; set; }
        string NewFieldName { set;  get; }
        bool CanBeReadWrite { set; get; }
        string FieldAccessExpression { get; }
        string AsTypeName { get; set; }
        string ParameterName { get;}
        bool ImplementLetSetterType { get; set; }
        bool ImplementSetSetterType { get; set; }
        bool FieldNameIsExemptFromValidation { get; }
        QualifiedModuleName QualifiedModuleName { get; }
        //DeclarationType DeclarationType { get; }
    }


    //Used for declarations that will be added to the code, but will never be encapsulated
    //Satifies the IFieldEncapsulationAttributes interface but some properties return n
    public class NeverEncapsulateAttributes : IFieldEncapsulationAttributes
    {
        private const string neverUse = "x_x_x_x";
        private IEncapsulateFieldNamesValidator _validator;
        private QualifiedModuleName _qmn;

        public NeverEncapsulateAttributes(string identifier, string asTypeName, QualifiedModuleName qmn, IEncapsulateFieldNamesValidator validator)
        {
            _qmn = qmn;
            _validator = validator;
            Identifier = identifier;
            NewFieldName = identifier;
            AsTypeName = asTypeName;
            FieldAccessExpressionFunc = () => NewFieldName;
        }

        public IFieldEncapsulationAttributes ApplyNewFieldName(string newFieldName)
        {
            NewFieldName = newFieldName;
            return this;
        }

        //public DeclarationType DeclarationType { get; } = DeclarationType.UserDefinedType;

        public string Identifier { private set; get; }

        public string NewFieldName { set; get; }

        string _tossString;
        public string PropertyName { set => _tossString = value; get => $"{neverUse}{Identifier}_{neverUse}"; }

        public string FieldAccessExpression => FieldAccessExpressionFunc();

        public Func<string> FieldAccessExpressionFunc { set; get; }

        public string AsTypeName { get; set; }
        public string ParameterName => neverUse;

        private bool _toss;
        public bool IsReadOnly { get; set; } = false;
        public bool CanBeReadWrite { get => false; set => _toss = value; }


        public bool EncapsulateFlag { get => false; set => _toss = value; }
        public bool ImplementLetSetterType { get => false; set => _toss = value; }
        public bool ImplementSetSetterType { get => false; set => _toss = value; }

        public bool FieldNameIsExemptFromValidation => false;
        public QualifiedModuleName QualifiedModuleName => _qmn;
    }

    public class FieldEncapsulationAttributes : IFieldEncapsulationAttributes
    {
        private QualifiedModuleName _qmn;
        private bool _fieldNameIsAlwaysValid;
        public FieldEncapsulationAttributes(Declaration target)
        {
            _fieldAndProperty = new EncapsulationIdentifiers(target);
            Identifier = target.IdentifierName;
            AsTypeName = target.AsTypeName;
            _qmn = target.QualifiedModuleName;
            FieldAccessExpressionFunc = () => NewFieldName;
            _fieldNameIsAlwaysValid = target.DeclarationType.Equals(DeclarationType.UserDefinedTypeMember);
            //DeclarationType = target.DeclarationType;
        }

        public FieldEncapsulationAttributes(string identifier, string asTypeName)
        {
            _fieldAndProperty = new EncapsulationIdentifiers(identifier);
            AsTypeName = asTypeName;
            FieldAccessExpressionFunc = () => NewFieldName;
        }

        public FieldEncapsulationAttributes(IFieldEncapsulationAttributes attributes)
        {
            _fieldAndProperty = new EncapsulationIdentifiers(attributes.Identifier, attributes.NewFieldName, attributes.PropertyName);
            PropertyName = attributes.PropertyName;
            IsReadOnly = attributes.IsReadOnly;
            EncapsulateFlag = attributes.EncapsulateFlag;
            NewFieldName = attributes.NewFieldName;
            CanBeReadWrite = attributes.CanBeReadWrite;
            AsTypeName = attributes.AsTypeName;
            ImplementLetSetterType = attributes.ImplementLetSetterType;
            ImplementSetSetterType = attributes.ImplementSetSetterType;
            QualifiedModuleName = attributes.QualifiedModuleName;
            FieldAccessExpressionFunc = () => NewFieldName;
        }

        private EncapsulationIdentifiers _fieldAndProperty;

        //public DeclarationType DeclarationType { private set; get; } = DeclarationType.Variable;

        public string Identifier { private set; get; }

        public bool CanBeReadWrite { set; get; } = true;

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

        public string FieldAccessExpression => FieldAccessExpressionFunc();

        public Func<string> FieldAccessExpressionFunc { set; get; }

        public string AsTypeName { get; set; }
        public string ParameterName => _fieldAndProperty.SetLetParameter;
        public bool IsReadOnly { get; set; }
        public bool EncapsulateFlag { get; set; }

        private bool _implLet;
        public bool ImplementLetSetterType { get => !IsReadOnly && _implLet; set => _implLet = value; }

        private bool _implSet;
        public bool ImplementSetSetterType { get => !IsReadOnly && _implSet; set => _implSet = value; }

        public bool FieldNameIsExemptFromValidation => _fieldNameIsAlwaysValid || NewFieldName.EqualsVBAIdentifier(Identifier);
        public QualifiedModuleName QualifiedModuleName
        {
            get => _qmn;
            set => _qmn = value;
        }
    }
}
