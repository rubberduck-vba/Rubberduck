using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;

namespace Rubberduck.Refactorings.EncapsulateField
{
                //    PropertyName = field.PropertyName,
                //AsTypeName = field.AsTypeName,
                //BackingField = field.PropertyAccessExpression(),
                //ParameterName = field.ParameterName,
                //GenerateSetter = field.ImplementSetSetterType,
                //GenerateLetter = field.ImplementLetSetterType

    //public interface IFieldEncapsulationAttributesX //: ISupportPropertyGenerator
    //{
    //    ////string IdentifierName { get; }
    //    ////string PropertyName { get; set; } //req'd
    //    ////bool IsReadOnly { get; set; }
    //    ////bool EncapsulateFlag { get; set; }
    //    ////string NewFieldName { set;  get; }
    //    ////bool CanBeReadWrite { set; get; }
    //    ////string AsTypeName { get; set; } //req'd
    //    ////string ParameterName { get; } //req'd
    //    ////bool ImplementLetSetterType { get; set; }//req'd
    //    ////bool ImplementSetSetterType { get; set; }//req'd
    //    ////bool FieldNameIsExemptFromValidation { get; }
    //    //QualifiedModuleName QualifiedModuleName { get; }
    //    ////Func<string> PropertyAccessExpression { set; get; } //req'd
    //    //Func<string> ReferenceExpression { set; get; }
    //}

    //public interface ISupportPropertyGenerator
    //{
    //    string PropertyName { get; set; } //req'd
    //    string AsTypeName { get; set; } //req'd
    //    string ParameterName { get; } //req'd
    //    bool ImplementLetSetterType { get; set; }//req'd
    //    bool ImplementSetSetterType { get; set; }//req'd
    //    Func<string> PropertyAccessExpression { set; get; } //req'd
    //}


    //Used for declarations that will be added to the code, but will never be encapsulated
    //Satifies the IFieldEncapsulationAttributes interface but some properties return n
    //public class NeverEncapsulateAttributes //: IFieldEncapsulationAttributes
    //{
    //    private const string neverUse = "x_x_x_x";
    //    private IEncapsulateFieldNamesValidator _validator;
    //    private QualifiedModuleName _qmn;

    //    public NeverEncapsulateAttributes(string identifier, string asTypeName, QualifiedModuleName qmn, IEncapsulateFieldNamesValidator validator)
    //    {
    //        _qmn = qmn;
    //        _validator = validator;
    //        IdentifierName = identifier;
    //        NewFieldName = identifier;
    //        AsTypeName = asTypeName;
    //        PropertyAccessExpression = () => NewFieldName;
    //        ReferenceExpression = () => NewFieldName;
    //    }

    //    //public IFieldEncapsulationAttributes ApplyNewFieldName(string newFieldName)
    //    //{
    //    //    NewFieldName = newFieldName;
    //    //    return this;
    //    //}

    //    public string IdentifierName { private set; get; }

    //    public string NewFieldName { set; get; }

    //    string _tossString;
    //    public string PropertyName { set => _tossString = value; get => $"{neverUse}{IdentifierName}_{neverUse}"; }

    //    public Func<string> PropertyAccessExpression { set; get; }

    //    public Func<string> ReferenceExpression { set; get; }

    //    public string AsTypeName { get; set; }
    //    public string ParameterName => neverUse;

    //    private bool _toss;
    //    public bool IsReadOnly { get; set; } = false;
    //    public bool CanBeReadWrite { get => false; set => _toss = value; }


    //    public bool EncapsulateFlag { get => false; set => _toss = value; }
    //    public bool ImplementLetSetterType { get => false; set => _toss = value; }
    //    public bool ImplementSetSetterType { get => false; set => _toss = value; }

    //    //public bool FieldNameIsExemptFromValidation => false;
    //    public QualifiedModuleName QualifiedModuleName => _qmn;
    //}

    //public class FieldEncapsulationAttributes //: IFieldEncapsulationAttributes
    //{
    //    private QualifiedModuleName _qmn;
    //    private bool _fieldNameIsAlwaysValid;
    //    public FieldEncapsulationAttributes(Declaration target)
    //    {
    //        _fieldAndProperty = new EncapsulationIdentifiers(target);
    //        IdentifierName = target.IdentifierName;
    //        AsTypeName = target.AsTypeName;
    //        _qmn = target.QualifiedModuleName;
    //        PropertyAccessExpression = () => NewFieldName;
    //        //ReferenceExpression = () => NewFieldName;
    //        ReferenceExpression = () => PropertyName;
    //        _fieldNameIsAlwaysValid = target.DeclarationType.Equals(DeclarationType.UserDefinedTypeMember);
    //    }

    //    //public FieldEncapsulationAttributes(string identifier, string asTypeName)
    //    //{
    //    //    _fieldAndProperty = new EncapsulationIdentifiers(identifier);
    //    //    AsTypeName = asTypeName;
    //    //    PropertyAccessExpression = () => NewFieldName;
    //    //    ReferenceExpression = () => PropertyName;
    //    //}

    //    //public FieldEncapsulationAttributes(IFieldEncapsulationAttributes attributes)
    //    //{
    //        //_fieldAndProperty = new EncapsulationIdentifiers(attributes.IdentifierName, attributes.NewFieldName, attributes.PropertyName);
    //        //PropertyName = attributes.PropertyName;
    //        //IsReadOnly = attributes.IsReadOnly;
    //        //EncapsulateFlag = attributes.EncapsulateFlag;
    //        //NewFieldName = attributes.NewFieldName;
    //        //CanBeReadWrite = attributes.CanBeReadWrite;
    //        //AsTypeName = attributes.AsTypeName;
    //        //ImplementLetSetterType = attributes.ImplementLetSetterType;
    //        //ImplementSetSetterType = attributes.ImplementSetSetterType;
    //        //QualifiedModuleName = attributes.QualifiedModuleName;
    //        //PropertyAccessExpression = () => NewFieldName;
    //        //ReferenceExpression = () => PropertyName;
    //    //}

    //    private EncapsulationIdentifiers _fieldAndProperty;

    //    public string IdentifierName { private set; get; }

    //    public bool CanBeReadWrite { set; get; } = true;

    //    public string NewFieldName
    //    {
    //        get => _fieldAndProperty.Field;
    //        set => _fieldAndProperty.Field = value;
    //    }

    //    public string PropertyName
    //    {
    //        get => _fieldAndProperty.Property;
    //        set => _fieldAndProperty.Property = value;
    //    }

    //    private Func<string> _propertyAccessExpression;
    //    public Func<string> PropertyAccessExpression
    //    {
    //        get
    //        {
    //            var test = _propertyAccessExpression();
    //            return _propertyAccessExpression;
    //        }
    //        set
    //        {
    //            _propertyAccessExpression = value;
    //            var test = value();
    //        }
    //    }

    //    public Func<string> ReferenceExpression { set; get; }

    //    public string AsTypeName { get; set; }
    //    public string ParameterName => _fieldAndProperty.SetLetParameter;
    //    public bool IsReadOnly { get; set; }
    //    public bool EncapsulateFlag { get; set; }

    //    private bool _implLet;
    //    public bool ImplementLetSetterType { get => !IsReadOnly && _implLet; set => _implLet = value; }

    //    private bool _implSet;
    //    public bool ImplementSetSetterType { get => !IsReadOnly && _implSet; set => _implSet = value; }

    //    public bool FieldNameIsExemptFromValidation => _fieldNameIsAlwaysValid || NewFieldName.EqualsVBAIdentifier(IdentifierName);
    //    public QualifiedModuleName QualifiedModuleName
    //    {
    //        get => _qmn;
    //        set => _qmn = value;
    //    }
    //}
}
