using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulatedFieldDeclaration
    {
        Declaration Declaration { get; }
        string IdentifierName { get; }
        DeclarationType DeclarationType { get; }
        IFieldEncapsulationAttributes EncapsulationAttributes { set; get; }
    }

    public class EncapsulatedFieldDeclaration : IEncapsulatedFieldDeclaration
    {
        private Declaration _decorated;
        public EncapsulatedFieldDeclaration(Declaration declaration)
        {
            _decorated = declaration;
            EncapsulationAttributes = new FieldEncapsulationAttributes(_decorated);
        }

        public Declaration Declaration => _decorated;

        public string IdentifierName => _decorated.IdentifierName;

        public DeclarationType DeclarationType => _decorated.DeclarationType;

        public IFieldEncapsulationAttributes EncapsulationAttributes { set; get; } = new FieldEncapsulationAttributes();

    }

    public class EncapsulatedValueType : EncapsulatedFieldDeclaration
    {
        public EncapsulatedValueType(Declaration declaration)
            : base(declaration)
        {
            EncapsulationAttributes.CanImplementLet = true;
            EncapsulationAttributes.CanImplementSet = false;
            EncapsulationAttributes.ImplementLetSetterType = true;
            EncapsulationAttributes.ImplementSetSetterType = false;
        }
    }

    public class EncapsulatedUserDefinedType : EncapsulatedFieldDeclaration
    {
        public EncapsulatedUserDefinedType(Declaration declaration)
            : base(declaration)
        {
            EncapsulationAttributes.CanImplementLet = true;
            EncapsulationAttributes.CanImplementSet = false;
            EncapsulationAttributes.ImplementLetSetterType = true;
            EncapsulationAttributes.ImplementSetSetterType = false;
        }
    }

    public class EncapsulatedVariantType : EncapsulatedFieldDeclaration
    {
        public EncapsulatedVariantType(Declaration declaration)
            : base(declaration)
        {
            EncapsulationAttributes.CanImplementLet = true;
            EncapsulationAttributes.CanImplementSet = true;
            EncapsulationAttributes.ImplementLetSetterType = true;
            EncapsulationAttributes.ImplementSetSetterType = true;
        }
    }

    public class EncapsulatedObjectType : EncapsulatedFieldDeclaration
    {
        public EncapsulatedObjectType(Declaration declaration)
            : base(declaration)
        {
            EncapsulationAttributes.CanImplementLet = false;
            EncapsulationAttributes.CanImplementSet = true;
            EncapsulationAttributes.ImplementLetSetterType = false;
            EncapsulationAttributes.ImplementSetSetterType = true;
        }
    }

    public class EncapsulatedArrayType : EncapsulatedFieldDeclaration
    {
        public EncapsulatedArrayType(Declaration declaration)
            : base(declaration)
        {
            EncapsulationAttributes.CanImplementLet = false;
            EncapsulationAttributes.CanImplementSet = false;
            EncapsulationAttributes.ImplementLetSetterType = false;
            EncapsulationAttributes.ImplementSetSetterType = false;
            EncapsulationAttributes.AsTypeName = Tokens.Variant;
        }
    }

    public class EncapsulatedUserDefinedTypeMember : IEncapsulatedFieldDeclaration
    {
        private readonly IEncapsulatedFieldDeclaration _decorated;
        public EncapsulatedUserDefinedTypeMember(IEncapsulatedFieldDeclaration encapsulateFieldDeclaration, EncapsulatedUserDefinedType udtVariable)
        {
            _decorated = encapsulateFieldDeclaration;
            EncapsulationAttributes = _decorated.EncapsulationAttributes;
            EncapsulationAttributes.NewFieldName = $"{udtVariable.Declaration.IdentifierName}.{encapsulateFieldDeclaration.Declaration.IdentifierName}";
        }

        public Declaration Declaration => _decorated.Declaration;

        public string IdentifierName => _decorated.IdentifierName;

        public DeclarationType DeclarationType => _decorated.DeclarationType;

        public IFieldEncapsulationAttributes EncapsulationAttributes { set; get; }
    }
}
