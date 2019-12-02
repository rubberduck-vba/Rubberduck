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
        DeclarationType DeclarationType { get; }
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
        string FieldReferenceExpression { get; }
        bool IsUDTMember { get; }
        bool HasValidEncapsulationAttributes { get; }
        QualifiedModuleName QualifiedModuleName { get; }
        IEnumerable<IdentifierReference> References { get; }
        bool ImplementLetSetterType { get; }
        bool ImplementSetSetterType { get; }
    }

    public class EncapsulateFieldCandidate : IEncapsulateFieldCandidate
    {
        protected Declaration _target;
        protected IFieldEncapsulationAttributes _attributes;
        protected FieldEncapsulationAttributes _concreteAttributes;
        private IEncapsulateFieldNamesValidator _validator;

        public EncapsulateFieldCandidate(Declaration declaration, IEncapsulateFieldNamesValidator validator)
        {
            _target = declaration;
            _concreteAttributes = new FieldEncapsulationAttributes(_target);
            _attributes = _concreteAttributes;
            _validator = validator;
            _validator.ForceNonConflictEncapsulationAttributes(_attributes, _target.QualifiedModuleName, _target);
        }

        public EncapsulateFieldCandidate(IFieldEncapsulationAttributes attributes, IEncapsulateFieldNamesValidator validator)
        {
            _target = null;
            _concreteAttributes = new FieldEncapsulationAttributes(attributes);
            _attributes = _concreteAttributes;
            _validator = validator;
            _validator.ForceNonConflictEncapsulationAttributes(_attributes, _attributes.QualifiedModuleName, _target);
        }

        //TODO: Defaulting to DeclarationType.Variable needs to be better
        private static void ForceNonConflictNewName(string identifier, QualifiedModuleName qmn, IFieldEncapsulationAttributes attributes, IEncapsulateFieldNamesValidator validator, IEnumerable<Declaration> ignore)
        {
            var isValidAttributeSet = validator.HasValidEncapsulationAttributes(attributes, qmn, ignore, DeclarationType.Variable);
            for (var idx = 1; idx < 9 && !isValidAttributeSet; idx++)
            {
                attributes.NewFieldName = $"{identifier}{idx}";
                isValidAttributeSet = validator.HasValidEncapsulationAttributes(attributes, qmn, ignore, DeclarationType.Variable);
            }
        }

        //TODO: Defaulting to DeclarationType.Variable needs to be better
        private static void ForceNonConflictPropertyName(string identifier, QualifiedModuleName qmn, IFieldEncapsulationAttributes attributes, IEncapsulateFieldNamesValidator validator, IEnumerable<Declaration> ignore)
        {
            var isValidAttributeSet = validator.HasValidEncapsulationAttributes(attributes, qmn, ignore, DeclarationType.Variable);
            for (var idx = 1; idx < 9 && !isValidAttributeSet; idx++)
            {
                attributes.PropertyName = $"{identifier}{idx}";
                isValidAttributeSet = validator.HasValidEncapsulationAttributes(attributes, qmn, ignore, DeclarationType.Variable);
            }
        }

        public Declaration Declaration => _target;

        public DeclarationType DeclarationType => _target?.DeclarationType ?? DeclarationType.Variable;

        public bool HasValidEncapsulationAttributes
        {
            get
            {
                var ignore = _target != null ? new Declaration[] { _target } : Enumerable.Empty<Declaration>();
                var declarationType = _target != null ? _target.DeclarationType : DeclarationType.Variable;
                return _validator.HasValidEncapsulationAttributes(EncapsulationAttributes, QualifiedModuleName, ignore, declarationType); //(Declaration dec) => dec.Equals(_target));
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

        public virtual bool IsUDTMember { get; } = false;

        public QualifiedModuleName QualifiedModuleName => Declaration?.QualifiedModuleName ?? _attributes.QualifiedModuleName;

        public IEnumerable<IdentifierReference> References => Declaration?.References ?? Enumerable.Empty<IdentifierReference>();

        public string IdentifierName => Declaration?.IdentifierName ?? _attributes.Identifier;

        public string ParameterName => _attributes.ParameterName;

        public string FieldReferenceExpression => _attributes.FieldAccessExpression;

        public bool ImplementLetSetterType { get => _attributes.ImplementLetSetterType; set => _attributes.ImplementLetSetterType = value; }
        public bool ImplementSetSetterType { get => _attributes.ImplementSetSetterType; set => _attributes.ImplementSetSetterType = value; }
    }

    public class EncapsulatedUserDefinedTypeField : EncapsulateFieldCandidate
    {
        public List<IEncapsulateFieldCandidate> Members { set; get; } = new List<IEncapsulateFieldCandidate>();
        public EncapsulatedUserDefinedTypeField(Declaration declaration, IEncapsulateFieldNamesValidator validator)
            : base(declaration, validator) { }
    }

    public interface IEncapsulatedUserDefinedTypeMember
    {
        Func<string> FieldAccessExpressionFunc { set; get; }
    }


    public class EncapsulatedUserDefinedTypeMember : EncapsulateFieldCandidate, IEncapsulatedUserDefinedTypeMember
    {
        public EncapsulatedUserDefinedTypeMember(Declaration target, IEncapsulateFieldCandidate udtVariable, IEncapsulateFieldNamesValidator validator, bool propertyNameRequiresParentIdentifier)
            : base(target, validator)
        {
            Parent = udtVariable;
            NameResolveProperty = propertyNameRequiresParentIdentifier;

            EncapsulationAttributes.PropertyName = NameResolveProperty
                ? $"{Parent.IdentifierName.Capitalize()}_{IdentifierName}"
                : IdentifierName;

            FieldAccessExpressionFunc =
                   () =>
                   {
                       var prefix = Parent.EncapsulateFlag
                                      ? Parent.NewFieldName
                                      : Parent.IdentifierName;

                       return $"{prefix}.{NewFieldName}";
                   };
        }

        public Func<string> FieldAccessExpressionFunc
        {
            set => _concreteAttributes.FieldAccessExpressionFunc = value;
            get => _concreteAttributes.FieldAccessExpressionFunc;
        }

        public IEncapsulateFieldCandidate Parent { private set; get; }

        public bool NameResolveProperty { set; get; }

        public override string TargetID => $"{Parent.IdentifierName}.{IdentifierName}";

        public override bool IsUDTMember => true;
    }
}
