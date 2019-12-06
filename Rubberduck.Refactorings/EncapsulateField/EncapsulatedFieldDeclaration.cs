using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldCandidate : IFieldEncapsulationAttributes
    {
        Declaration Declaration { get; }
        string TargetID { get; }
        //string IdentifierName { get; }
        IFieldEncapsulationAttributes EncapsulationAttributes { get; }
        //bool IsReadOnly { set; get; }
        //bool CanBeReadWrite { get; }
        //string PropertyName { set; get; }
        //bool EncapsulateFlag { set; get; }
        //string NewFieldName { get; }
        //string AsTypeName { get; }
        //string ParameterName { get; }
        bool IsUDTMember { get; }
        bool HasValidEncapsulationAttributes { get; }
        //QualifiedModuleName QualifiedModuleName { get; }
        IEnumerable<IdentifierReference> References { get; }
        //bool ImplementLetSetterType { get; }
        //bool ImplementSetSetterType { get; }
        //Func<string> PropertyAccessExpression { set; get; }
        //Func<string> ReferenceExpression { set; get; }
        string this[IdentifierReference idRef] { set; get; }
        bool TryGetReferenceExpression(IdentifierReference idRef, out string expression);
        bool FieldNameIsExemptFromValidation { get; }
    }

    public interface IEncapsulatedUserDefinedTypeField : IEncapsulateFieldCandidate
    {
        IList<IEncapsulatedUserDefinedTypeMember> Members { set; get; }
        bool TypeDeclarationIsPrivate { set; get; }
    }

    public class EncapsulateFieldCandidate : IEncapsulateFieldCandidate
    {
        protected Declaration _target;
        protected IFieldEncapsulationAttributes _attributes;
        private IEncapsulateFieldNamesValidator _validator;
        private Dictionary<IdentifierReference, string> _idRefRenames { set; get; } = new Dictionary<IdentifierReference, string>();

        public EncapsulateFieldCandidate(Declaration declaration, IEncapsulateFieldNamesValidator validator)
        {
            _target = declaration;
            AsTypeName = _target.AsTypeName;
            _attributes = new FieldEncapsulationAttributes(_target);
            _validator = validator;
            _validator.ForceNonConflictEncapsulationAttributes(this, _target.QualifiedModuleName, _target);
        }

        public EncapsulateFieldCandidate(IFieldEncapsulationAttributes attributes, IEncapsulateFieldNamesValidator validator)
        {
            _target = null;
            _attributes = new FieldEncapsulationAttributes(attributes);
            _validator = validator;
            _validator.ForceNonConflictEncapsulationAttributes(this, _attributes.QualifiedModuleName, _target);
        }

        public Declaration Declaration => _target;

        public bool HasValidEncapsulationAttributes
        {
            get
            {
                var ignore = _target != null ? new Declaration[] { _target } : Enumerable.Empty<Declaration>();
                var declarationType = _target != null ? _target.DeclarationType : DeclarationType.Variable;
                return _validator.HasValidEncapsulationAttributes(this, QualifiedModuleName, ignore, declarationType); //(Declaration dec) => dec.Equals(_target));
            }
        }

        public bool TryGetReferenceExpression(IdentifierReference idRef, out string expression)
        {
            expression = string.Empty;
            if (_idRefRenames.ContainsKey(idRef))
            {
                expression = _idRefRenames[idRef];
            }
            return expression.Length > 0;
        }

        public string this[IdentifierReference idRef]
        {
            get => _idRefRenames[idRef];
            set
            {
                if (!_idRefRenames.ContainsKey(idRef))
                {
                    _idRefRenames.Add(idRef, value);
                }
                else
                {
                    _idRefRenames[idRef] = value;
                }
            }
        }


        public IFieldEncapsulationAttributes EncapsulationAttributes
        {
            private set => _attributes = value;
            get => this as IFieldEncapsulationAttributes; // _attributes;
        }

        public virtual string TargetID => _target?.IdentifierName ?? _attributes.IdentifierName;

        public bool EncapsulateFlag
        {
            get => _attributes.EncapsulateFlag;
            set => _attributes.EncapsulateFlag = value;
        }

        public bool IsReadOnly
        {
            get => _attributes.IsReadOnly;
            set => _attributes.IsReadOnly = value;
        }

        public bool CanBeReadWrite
        {
            get => _attributes.CanBeReadWrite;
            set => _attributes.CanBeReadWrite = value;
        }

        public string PropertyName
        {
            get => _attributes.PropertyName;
            set => _attributes.PropertyName = value;
        }

        public bool IsEditableReadWriteFieldIdentifier { set; get; } = true;

        public string NewFieldName
        {
            get => _attributes.NewFieldName;
            set => _attributes.NewFieldName = value;
        }

        private string _asTypeName;
        public string AsTypeName
        {
            set
            {
                _asTypeName = value;
                if (_attributes != null) { _attributes.AsTypeName = value; }
            }
            get
            {
                return _attributes?.AsTypeName ?? _asTypeName;
            }
        } //=> _target?.AsTypeName ?? _attributes.AsTypeName;
        //{
        //    get => /*_target?.AsTypeName ??*/ _attributes.AsTypeName;
        //    set => _attributes.AsTypeName = value;
        //    //{
        //    //    if (_target is null) { _attributes.AsTypeName = value; }
        //    //}
        //}

        public bool IsUDTMember => _target?.DeclarationType.Equals(DeclarationType.UserDefinedTypeMember) ?? false;

        public QualifiedModuleName QualifiedModuleName => Declaration?.QualifiedModuleName ?? _attributes.QualifiedModuleName;

        public IEnumerable<IdentifierReference> References => Declaration?.References ?? Enumerable.Empty<IdentifierReference>();

        public string IdentifierName
        {
            get => Declaration?.IdentifierName ?? _attributes.IdentifierName;
        }

        public string ParameterName => _attributes.ParameterName;

        public bool ImplementLetSetterType { get => _attributes.ImplementLetSetterType; set => _attributes.ImplementLetSetterType = value; }
        public bool ImplementSetSetterType { get => _attributes.ImplementSetSetterType; set => _attributes.ImplementSetSetterType = value; }

        public Func<string> PropertyAccessExpression
        {
            //set => _attributes.PropertyAccessExpression = value;
            get => _attributes.PropertyAccessExpression;
            set
            {
                _attributes.PropertyAccessExpression = value;
                var test = value();
            }
        }

        public Func<string> ReferenceExpression
        {
            get => _attributes.ReferenceExpression;
            set
            {
                _attributes.ReferenceExpression = value;
                var test = value();
            }
        }

        public bool FieldNameIsExemptFromValidation 
            => Declaration?.DeclarationType.Equals(DeclarationType.UserDefinedTypeMember) ?? false || NewFieldName.EqualsVBAIdentifier(IdentifierName);
    }

    public class EncapsulatedUserDefinedTypeField : EncapsulateFieldCandidate, IEncapsulatedUserDefinedTypeField
    {
        public EncapsulatedUserDefinedTypeField(Declaration declaration, IEncapsulateFieldNamesValidator validator)
            : base(declaration, validator)
        {
            PropertyAccessExpression = () => EncapsulateFlag ? NewFieldName : IdentifierName;
        }

        public IList<IEncapsulatedUserDefinedTypeMember> Members { set; get; } = new List<IEncapsulatedUserDefinedTypeMember>();
        public bool TypeDeclarationIsPrivate { set; get; }
    }

    public interface IEncapsulatedUserDefinedTypeMember : IEncapsulateFieldCandidate
    {
        IEncapsulateFieldCandidate Parent { get; }
        bool FieldQualifyProperty { set; get; }
    }

    public class EncapsulatedUserDefinedTypeMember : EncapsulateFieldCandidate, IEncapsulatedUserDefinedTypeMember
    {
        public EncapsulatedUserDefinedTypeMember(Declaration target, IEncapsulateFieldCandidate udtVariable, IEncapsulateFieldNamesValidator validator)
            : base(target, validator)
        {
            Parent = udtVariable;

            PropertyName = IdentifierName;
            PropertyAccessExpression = () => $"{Parent.PropertyAccessExpression()}.{PropertyName}";
            ReferenceExpression = () => $"{Parent.PropertyAccessExpression()}.{PropertyName}";
        }

        public IEncapsulateFieldCandidate Parent { private set; get; }

        private bool _fieldNameQualifyProperty;
        public bool FieldQualifyProperty
        {
            get => _fieldNameQualifyProperty;
            set
            {
                _fieldNameQualifyProperty = value;
                PropertyName = _fieldNameQualifyProperty
                    ? $"{Parent.IdentifierName.Capitalize()}_{IdentifierName}"
                    : IdentifierName;
            }
        }
        public override string TargetID => $"{Parent.IdentifierName}.{IdentifierName}";
    }
}
