using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulatedFieldDeclaration
    {
        Declaration Declaration { get; }
        string IdentifierName { get; }
        DeclarationType DeclarationType { get; }
        Accessibility Accessibility { get;}
        IFieldEncapsulationAttributes EncapsulationAttributes { set; get; }
    }

    public class EncapsulatedFieldDeclaration : IEncapsulatedFieldDeclaration, IFieldEncapsulationAttributes
    {
        protected Declaration _decorated;
        protected IFieldEncapsulationAttributes _attributes;

        public EncapsulatedFieldDeclaration(Declaration declaration)
        {
            _decorated = declaration;
            _attributes = new FieldEncapsulationAttributes(_decorated);
        }

        public Declaration Declaration => _decorated;

        public string IdentifierName => _decorated.IdentifierName;

        public DeclarationType DeclarationType => _decorated.DeclarationType;

        public Accessibility Accessibility => _decorated.Accessibility;

        public IFieldEncapsulationAttributes EncapsulationAttributes
        {
            set => _attributes = value;
            get => _attributes;
        }

        public string AsTypeName { get => _attributes.AsTypeName; set => _attributes.AsTypeName = value; }
        public string ParameterName { get => _attributes.ParameterName; set => _attributes.ParameterName = value; }
        public bool CanImplementLet { get => _attributes.CanImplementLet; set => _attributes.CanImplementLet = value; }
        public bool CanImplementSet { get => _attributes.CanImplementSet; set => _attributes.CanImplementSet = value; }
        public string FieldName { get => _attributes.FieldName; set => _attributes.FieldName = value; }
        public string NewFieldName { get => _attributes.NewFieldName; set => _attributes.NewFieldName = value; }
        public string PropertyName { get => _attributes.PropertyName; set => _attributes.PropertyName = value; }
        public bool ImplementLetSetterType { get => _attributes.ImplementLetSetterType; set => _attributes.ImplementLetSetterType = value; }
        public bool ImplementSetSetterType { get => _attributes.ImplementSetSetterType; set => _attributes.ImplementSetSetterType = value; }
        public bool EncapsulateFlag { get => _attributes.EncapsulateFlag; set => _attributes.EncapsulateFlag = value; }
    }

    public class EncapsulateFieldDecoratorBase : IEncapsulatedFieldDeclaration, IFieldEncapsulationAttributes
    {

        private IEncapsulatedFieldDeclaration _decorated;
        private IFieldEncapsulationAttributes _attributes;
        public EncapsulateFieldDecoratorBase(IEncapsulatedFieldDeclaration efd)
        {
            _decorated = efd;
            _attributes = efd as IFieldEncapsulationAttributes;
        }

        public Declaration Declaration => _decorated.Declaration;

        public string IdentifierName => _decorated.IdentifierName;

        public DeclarationType DeclarationType => _decorated.DeclarationType;

        public Accessibility Accessibility => _decorated.Accessibility;

        public IFieldEncapsulationAttributes EncapsulationAttributes { get => _decorated.EncapsulationAttributes; set => _decorated.EncapsulationAttributes = value; }
        public string AsTypeName { get => _attributes.AsTypeName; set => _attributes.AsTypeName = value; }
        public string ParameterName { get => _attributes.ParameterName; set => _attributes.ParameterName = value; }
        public bool CanImplementLet { get => _attributes.CanImplementLet; set => _attributes.CanImplementLet = value; }
        public bool CanImplementSet { get => _attributes.CanImplementSet; set => _attributes.CanImplementSet = value; }
        public string FieldName { get => _attributes.FieldName; set => _attributes.FieldName = value; }
        public string NewFieldName { get => _attributes.NewFieldName; set => _attributes.NewFieldName = value; }
        public string PropertyName { get => _attributes.PropertyName; set => _attributes.PropertyName = value; }
        public bool ImplementLetSetterType { get => _attributes.ImplementLetSetterType; set => _attributes.ImplementLetSetterType = value; }
        public bool ImplementSetSetterType { get => _attributes.ImplementSetSetterType; set => _attributes.ImplementSetSetterType = value; }
        public bool EncapsulateFlag { get => _attributes.EncapsulateFlag; set => _attributes.EncapsulateFlag = value; }
    }

    public class EncapsulatedValueType : EncapsulateFieldDecoratorBase // IEncapsulatedFieldDeclaration, IFieldEncapsulationAttributes
    {
        private EncapsulatedValueType(IEncapsulatedFieldDeclaration efd)
            : base(efd)
        {
            CanImplementLet = true;
            CanImplementSet = false;
            ImplementLetSetterType = true;
            ImplementSetSetterType = false;
        }

        public static IEncapsulatedFieldDeclaration Decorate(IEncapsulatedFieldDeclaration efd) 
            => new EncapsulatedValueType(efd);
    }

    public class EncapsulatedUserDefinedType : EncapsulateFieldDecoratorBase
    {

        private EncapsulatedUserDefinedType(IEncapsulatedFieldDeclaration efd)
        : base(efd)
        {
            CanImplementLet = true;
            CanImplementSet = false;
            ImplementLetSetterType = true;
            ImplementSetSetterType = false;
        }

        public static IEncapsulatedFieldDeclaration Decorate(IEncapsulatedFieldDeclaration efd) 
            => new EncapsulatedUserDefinedType(efd);
    }

    public class EncapsulatedVariantType : EncapsulateFieldDecoratorBase // IEncapsulatedFieldDeclaration, IFieldEncapsulationAttributes
    {
        private EncapsulatedVariantType(IEncapsulatedFieldDeclaration efd)
        : base(efd)
        {
            CanImplementLet = true;
            CanImplementSet = true;
            ImplementLetSetterType = true;
            ImplementSetSetterType = true;
        }

        public static IEncapsulatedFieldDeclaration Decorate(IEncapsulatedFieldDeclaration efd) 
            => new EncapsulatedVariantType(efd);
    }

    public class EncapsulatedObjectType : EncapsulateFieldDecoratorBase //IEncapsulatedFieldDeclaration, IFieldEncapsulationAttributes
    {
        private EncapsulatedObjectType(IEncapsulatedFieldDeclaration efd)
            : base(efd)
        {
            CanImplementLet = false;
            CanImplementSet = true;
            ImplementLetSetterType = false;
            ImplementSetSetterType = true;
        }

        public static IEncapsulatedFieldDeclaration Decorate(IEncapsulatedFieldDeclaration efd) 
            => new EncapsulatedObjectType(efd);
    }

    public class EncapsulatedArrayType : EncapsulateFieldDecoratorBase //IEncapsulatedFieldDeclaration, IFieldEncapsulationAttributes
    {
        private EncapsulatedArrayType(IEncapsulatedFieldDeclaration efd)
        : base(efd)
        {
            CanImplementLet = false;
            CanImplementSet = false;
            ImplementLetSetterType = false;
            ImplementSetSetterType = false;
            AsTypeName = Tokens.Variant;
        }

        public static IEncapsulatedFieldDeclaration Decorate(IEncapsulatedFieldDeclaration efd) 
            => new EncapsulatedArrayType(efd);
    }

    public class EncapsulatedUserDefinedTypeMember : EncapsulateFieldDecoratorBase
    {
        private EncapsulatedUserDefinedType _udt;

        private EncapsulatedUserDefinedTypeMember(IEncapsulatedFieldDeclaration efd, EncapsulatedUserDefinedType udtVariable)
        : base(efd)
        {
            _udt = udtVariable;
            var decoratedAttributes = efd as IFieldEncapsulationAttributes;
            CanImplementLet = decoratedAttributes.CanImplementLet;
            CanImplementSet = decoratedAttributes.CanImplementSet;
            ImplementLetSetterType = decoratedAttributes.ImplementLetSetterType;
            ImplementSetSetterType = decoratedAttributes.ImplementSetSetterType;
            NewFieldName = $"{udtVariable.Declaration.IdentifierName}.{efd.Declaration.IdentifierName}";
        }

        public static IEncapsulatedFieldDeclaration Decorate(IEncapsulatedFieldDeclaration efd, EncapsulatedUserDefinedType udtVariable) 
            => new EncapsulatedUserDefinedTypeMember(efd, udtVariable);
    }
}
