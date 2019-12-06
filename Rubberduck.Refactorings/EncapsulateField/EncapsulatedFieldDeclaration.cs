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
        string IdentifierName { get; }
        string TargetID { get; }
        bool IsReadOnly { get; set; }
        bool EncapsulateFlag { get; set; }
        string NewFieldName { set; get; }
        bool CanBeReadWrite { set; get; }
        bool IsUDTMember { get; }
        bool HasValidEncapsulationAttributes { get; }
        QualifiedModuleName QualifiedModuleName { get; }
        IEnumerable<IdentifierReference> References { get; }
        string this[IdentifierReference idRef] { set; get; }
        bool TryGetReferenceExpression(IdentifierReference idRef, out string expression);
        bool FieldNameIsExemptFromValidation { get; }
        Func<string> ReferenceExpression { set; get; }
        string PropertyName { get; set; }
        string AsTypeName { get; set; } 
        string ParameterName { get; } 
        bool ImplementLetSetterType { get; set; }
        bool ImplementSetSetterType { get; set; }
        Func<string> PropertyAccessExpression { set; get; }
    }

    public interface IEncapsulatedUserDefinedTypeField : IEncapsulateFieldCandidate, ISupportPropertyGenerator
    {
        IList<IEncapsulatedUserDefinedTypeMember> Members { set; get; }
        bool TypeDeclarationIsPrivate { set; get; }
    }

    public class EncapsulateFieldCandidate : IEncapsulateFieldCandidate, ISupportPropertyGenerator
    {
        protected Declaration _target;
        protected QualifiedModuleName _qmn;
        private IEncapsulateFieldNamesValidator _validator;
        private Dictionary<IdentifierReference, string> _idRefRenames { set; get; } = new Dictionary<IdentifierReference, string>();
        private EncapsulationIdentifiers _fieldAndProperty;

        public EncapsulateFieldCandidate(Declaration declaration, IEncapsulateFieldNamesValidator validator)
        {
            _target = declaration;
            AsTypeName = _target.AsTypeName;

            _fieldAndProperty = new EncapsulationIdentifiers(_target);
            IdentifierName = _target.IdentifierName;
            AsTypeName = _target.AsTypeName;
            _qmn = _target.QualifiedModuleName;
            PropertyAccessExpression = () => NewFieldName;
            ReferenceExpression = () => PropertyName;

            _validator = validator;
            _validator.ForceNonConflictEncapsulationAttributes(this, _target.QualifiedModuleName, _target);
        }

        public EncapsulateFieldCandidate(string identifier, string asTypeName, QualifiedModuleName qmn,/*IFieldEncapsulationAttributes attributes,*/ IEncapsulateFieldNamesValidator validator, bool neverEncapsulate = false)
        {
            _target = null;

            _fieldAndProperty = new EncapsulationIdentifiers(identifier, neverEncapsulate);
            IdentifierName = identifier;
            AsTypeName = asTypeName;
            _qmn = qmn;
            PropertyAccessExpression = () => NewFieldName;
            ReferenceExpression = () => PropertyName;

            _validator = validator;
            _validator.ForceNonConflictEncapsulationAttributes(this, qmn, _target);
        }

        public Declaration Declaration => _target;

        public bool HasValidEncapsulationAttributes
        {
            get
            {
                var ignore = _target != null ? new Declaration[] { _target } : Enumerable.Empty<Declaration>();
                var declarationType = _target != null ? _target.DeclarationType : DeclarationType.Variable;
                return _validator.HasValidEncapsulationAttributes(this, QualifiedModuleName, ignore, declarationType);
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

        public virtual string TargetID => _target?.IdentifierName ?? IdentifierName;

        public bool EncapsulateFlag { set; get; }
        public bool IsReadOnly { set; get; }
        public bool CanBeReadWrite { set; get; }

        public string PropertyName
        {
            get => _fieldAndProperty.Property;
            set => _fieldAndProperty.Property = value;
        }

        public bool IsEditableReadWriteFieldIdentifier { set; get; } = true;

        public string NewFieldName
        {
            get => _fieldAndProperty.Field;
            set => _fieldAndProperty.Field = value;
        }

        public string AsTypeName { set; get; }

        public bool IsUDTMember => _target?.DeclarationType.Equals(DeclarationType.UserDefinedTypeMember) ?? false;

        public QualifiedModuleName QualifiedModuleName => _qmn;

        public IEnumerable<IdentifierReference> References => Declaration?.References ?? Enumerable.Empty<IdentifierReference>();

        private string _identifierName;
        public string IdentifierName
        {
            get => Declaration?.IdentifierName ?? _identifierName;
            set => _identifierName = value;
        }

        public string ParameterName => _fieldAndProperty.SetLetParameter;

        private bool _implLet;
        public bool ImplementLetSetterType { get => !IsReadOnly && _implLet; set => _implLet = value; }

        private bool _implSet;
        public bool ImplementSetSetterType { get => !IsReadOnly && _implSet; set => _implSet = value; }

        public Func<string> PropertyAccessExpression { set; get; }

        public Func<string> ReferenceExpression { set; get; }

        public bool FieldNameIsExemptFromValidation 
            => Declaration?.DeclarationType.Equals(DeclarationType.UserDefinedTypeMember) ?? false;
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
