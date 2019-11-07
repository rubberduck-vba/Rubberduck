using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldDeclaration
    {
        private Declaration _decorated;
        public EncapsulateFieldDeclaration(Declaration declaration)
        {
            _decorated = declaration;
            EncapsulationAttributes = new EncapsulationAttributes(_decorated);
        }

        public static EncapsulateFieldDeclaration Decorate(Declaration declaration)
        {
            return new EncapsulateFieldDeclaration(declaration);
        }

        public Declaration Declaration => _decorated;

        public string IdentifierName => _decorated.IdentifierName;

        public DeclarationType DeclarationType => _decorated.DeclarationType;

        public EncapsulationAttributes EncapsulationAttributes { set; get; }
    }

    public class EncapsulatedValueType : EncapsulateFieldDeclaration
    {
        public EncapsulatedValueType(EncapsulateFieldDeclaration declaration)
            : base(declaration.Declaration)
        {
        }
    }

    public class EncapsulatedUserDefinedMemberValueType : EncapsulatedValueType
    {
        public EncapsulatedUserDefinedMemberValueType(EncapsulatedValueType declaration, UserDefinedTypeField udtVariable)
            : base(declaration)
        {
            EncapsulationAttributes.FieldName = $"{udtVariable.Declaration.IdentifierName}.{declaration.Declaration.IdentifierName}";
            EncapsulationAttributes.ImplementLetSetterType = true;
        }
    }

    public class UserDefinedTypeField : EncapsulateFieldDeclaration
    {
        public UserDefinedTypeField(Declaration declaration)
            : base(declaration)
        {
        }

        public UserDefinedTypeField(EncapsulateFieldDeclaration declaration)
            : base(declaration.Declaration)
        {
        }
    }

    public class EncapsulatedVariantType : EncapsulateFieldDeclaration
    {
        public EncapsulatedVariantType(EncapsulateFieldDeclaration declaration)
            : base(declaration.Declaration)
        {
            EncapsulationAttributes.CanImplementLet = true;
            EncapsulationAttributes.CanImplementSet = true;
        }
    }

    public class EncapsulatedObjectType : EncapsulateFieldDeclaration
    {
        public EncapsulatedObjectType(EncapsulateFieldDeclaration declaration)
            : base(declaration.Declaration)
        {
            EncapsulationAttributes.CanImplementLet = false;
            EncapsulationAttributes.CanImplementSet = true;
        }
    }
}
