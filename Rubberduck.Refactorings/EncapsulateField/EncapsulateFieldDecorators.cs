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
            EncapsulationAttributes = new FieldEncapsulationAttributes(_decorated);
        }

        public Declaration Declaration => _decorated;

        public string IdentifierName => _decorated.IdentifierName;

        public DeclarationType DeclarationType => _decorated.DeclarationType;

        public IFieldEncapsulationAttributes EncapsulationAttributes { set; get; }
    }

    public class EncapsulateValueType : EncapsulateFieldDeclaration
    {
        public EncapsulateValueType(EncapsulateFieldDeclaration declaration)
            : base(declaration.Declaration)
        {
            EncapsulationAttributes.CanImplementLet = true;
            EncapsulationAttributes.CanImplementSet = false;
        }
    }

    public class EncapsulatedUserDefinedMemberValueType : EncapsulateValueType
    {
        public EncapsulatedUserDefinedMemberValueType(EncapsulateValueType declaration, EncapsulateUserDefinedType udtVariable)
            : base(declaration)
        {
            EncapsulationAttributes.NewFieldName = $"{udtVariable.Declaration.IdentifierName}.{declaration.Declaration.IdentifierName}";
            EncapsulationAttributes.IsValueType = true;
            EncapsulationAttributes.CanImplementLet = true;
            EncapsulationAttributes.CanImplementSet = false;
        }

        public EncapsulatedUserDefinedMemberValueType(EncapsulateVariantType declaration, EncapsulateUserDefinedType udtVariable)
            : base(declaration)
        {
            EncapsulationAttributes.NewFieldName = $"{udtVariable.Declaration.IdentifierName}.{declaration.Declaration.IdentifierName}";
            EncapsulationAttributes.IsVariantType = true;
            EncapsulationAttributes.CanImplementLet = true;
            EncapsulationAttributes.CanImplementSet = false;
        }
    }

    public class EncapsulatedUserDefinedMemberObjectType : EncapsulateFieldDeclaration
    {
        public EncapsulatedUserDefinedMemberObjectType(EncapsulateObjectType declaration, EncapsulateUserDefinedType udtVariable)
            : base(declaration.Declaration)
        {
            EncapsulationAttributes.NewFieldName = $"{udtVariable.Declaration.IdentifierName}.{declaration.Declaration.IdentifierName}";
            EncapsulationAttributes.IsObjectType = true;
            EncapsulationAttributes.CanImplementLet = false;
            EncapsulationAttributes.CanImplementSet = true;
        }
    }


    public class EncapsulateUserDefinedType : EncapsulateFieldDeclaration
    {
        public EncapsulateUserDefinedType(Declaration declaration)
            : base(declaration)
        {
            EncapsulationAttributes.CanImplementLet = true;
            EncapsulationAttributes.CanImplementSet = false;
        }

        public EncapsulateUserDefinedType(EncapsulateFieldDeclaration declaration)
            : base(declaration.Declaration)
        {
            EncapsulationAttributes.CanImplementLet = true;
            EncapsulationAttributes.CanImplementSet = false;
        }
    }

    public class EncapsulateVariantType : EncapsulateFieldDeclaration
    {
        public EncapsulateVariantType(EncapsulateFieldDeclaration declaration)
            : base(declaration.Declaration)
        {
            EncapsulationAttributes.CanImplementLet = true;
            EncapsulationAttributes.CanImplementSet = true;
        }
        public EncapsulateVariantType(Declaration declaration)
            : base(declaration)
        {
            EncapsulationAttributes.CanImplementLet = true;
            EncapsulationAttributes.CanImplementSet = true;
        }
    }

    public class EncapsulateObjectType : EncapsulateFieldDeclaration
    {
        public EncapsulateObjectType(EncapsulateFieldDeclaration declaration)
            : base(declaration.Declaration)
        {
            EncapsulationAttributes.CanImplementLet = false;
            EncapsulationAttributes.CanImplementSet = true;
        }
    }
}
