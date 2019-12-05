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
    public interface IEncapsulateFieldCandidate
    {
        Declaration Declaration { get; }
        //DeclarationType DeclarationType { get; }
        string TargetID { get; }
        string IdentifierName { get; }
        IFieldEncapsulationAttributes EncapsulationAttributes { get; }
        bool IsReadOnly { set; get; }
        bool CanBeReadWrite { get; }
        string PropertyName { set; get; }
        bool EncapsulateFlag { set; get; }
        string NewFieldName { get; }
        string AsTypeName { get; }
        string ParameterName { get; }
        bool IsUDTMember { get; }
        bool HasValidEncapsulationAttributes { get; }
        QualifiedModuleName QualifiedModuleName { get; }
        IEnumerable<IdentifierReference> References { get; }
        bool ImplementLetSetterType { get; }
        bool ImplementSetSetterType { get; }
        Func<string> PropertyAccessExpression { set; get; }
        Func<string> ReferenceExpression { set; get; }
        string this[IdentifierReference idRef] { set; get; }
        bool TryGetReferenceExpression(IdentifierReference idRef, out string expression);
    }

    public interface IEncapsulatedUserDefinedTypeField : IEncapsulateFieldCandidate
    {
        IList<IEncapsulateFieldCandidate> Members { set; get; }
        IEnumerable<IEncapsulateFieldCandidate> SelectedMembers { get; }
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
            _attributes = new FieldEncapsulationAttributes(_target);
            _validator = validator;
            _validator.ForceNonConflictEncapsulationAttributes(_attributes, _target.QualifiedModuleName, _target);
        }

        public EncapsulateFieldCandidate(IFieldEncapsulationAttributes attributes, IEncapsulateFieldNamesValidator validator)
        {
            _target = null;
            _attributes = new FieldEncapsulationAttributes(attributes);
            _validator = validator;
            _validator.ForceNonConflictEncapsulationAttributes(_attributes, _attributes.QualifiedModuleName, _target);
        }

        public Declaration Declaration => _target;

        //public DeclarationType DeclarationType => _target?.DeclarationType ?? DeclarationType.Variable;

        public bool HasValidEncapsulationAttributes
        {
            get
            {
                var ignore = _target != null ? new Declaration[] { _target } : Enumerable.Empty<Declaration>();
                var declarationType = _target != null ? _target.DeclarationType : DeclarationType.Variable;
                return _validator.HasValidEncapsulationAttributes(EncapsulationAttributes, QualifiedModuleName, ignore, declarationType); //(Declaration dec) => dec.Equals(_target));
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
            get => _attributes;
        }

        public virtual string TargetID => _target?.IdentifierName ?? _attributes.Identifier;

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
        }

        public string AsTypeName => _target?.AsTypeName ?? _attributes.AsTypeName;

        public bool IsUDTMember => _target?.DeclarationType.Equals(DeclarationType.UserDefinedTypeMember) ?? false;

        public QualifiedModuleName QualifiedModuleName => Declaration?.QualifiedModuleName ?? _attributes.QualifiedModuleName;

        public IEnumerable<IdentifierReference> References => Declaration?.References ?? Enumerable.Empty<IdentifierReference>();

        public string IdentifierName => Declaration?.IdentifierName ?? _attributes.Identifier;

        public string ParameterName => _attributes.ParameterName;

        public bool ImplementLetSetterType { get => _attributes.ImplementLetSetterType; set => _attributes.ImplementLetSetterType = value; }
        public bool ImplementSetSetterType { get => _attributes.ImplementSetSetterType; set => _attributes.ImplementSetSetterType = value; }

        public Func<string> PropertyAccessExpression
        {
            set => _attributes.PropertyAccessExpression = value;
            get => _attributes.PropertyAccessExpression;
        }

        public Func<string> ReferenceExpression
        {
            set => _attributes.ReferenceExpression = value;
            get => _attributes.ReferenceExpression;
        }
    }

    public class EncapsulatedUserDefinedTypeField : EncapsulateFieldCandidate, IEncapsulatedUserDefinedTypeField
    {
        public EncapsulatedUserDefinedTypeField(Declaration declaration, IEncapsulateFieldNamesValidator validator)
            : base(declaration, validator)
        {
            PropertyAccessExpression = () => EncapsulateFlag ? NewFieldName : IdentifierName;
        }

        public IList<IEncapsulateFieldCandidate> Members { set; get; } = new List<IEncapsulateFieldCandidate>();
        public IEnumerable<IEncapsulateFieldCandidate> SelectedMembers => Members.Where(m => m.EncapsulateFlag);
        public bool TypeDeclarationIsPrivate { set; get; }
    }

    public class EncapsulatedUserDefinedTypeMember : EncapsulateFieldCandidate
    {
        public EncapsulatedUserDefinedTypeMember(Declaration target, IEncapsulateFieldCandidate udtVariable, IEncapsulateFieldNamesValidator validator, bool propertyNameRequiresParentIdentifier)
            : base(target, validator)
        {
            Parent = udtVariable;
            FieldNameQualifyProperty = propertyNameRequiresParentIdentifier;

            EncapsulationAttributes.PropertyName = FieldNameQualifyProperty
                ? $"{Parent.IdentifierName.Capitalize()}_{IdentifierName}"
                : IdentifierName;
        }

        public IEncapsulateFieldCandidate Parent { private set; get; }

        public bool FieldNameQualifyProperty { set; get; }

        public override string TargetID => $"{Parent.IdentifierName}.{IdentifierName}";
    }
}
