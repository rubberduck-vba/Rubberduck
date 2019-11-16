using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulatedFieldDeclaration
    {
        Declaration Declaration { get; }
        string IdentifierName { get; }
        KeyValuePair<Declaration, string> TargetIDPair { set;  get; }
        DeclarationType DeclarationType { get; }
        Accessibility Accessibility { get;}
        IFieldEncapsulationAttributes EncapsulationAttributes { set; get; }
    }

    public class EncapsulatedFieldDeclaration : IEncapsulatedFieldDeclaration
    {
        protected Declaration _decorated;
        private IFieldEncapsulationAttributes _attributes;

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

        public KeyValuePair<Declaration, string> TargetIDPair { get; set; }
    }

    public class EncapsulateFieldDecoratorBase : IEncapsulatedFieldDeclaration
    {

        private IEncapsulatedFieldDeclaration _decorated;

        public EncapsulateFieldDecoratorBase(IEncapsulatedFieldDeclaration efd)
        {
            _decorated = efd;
            TargetIDPair = new KeyValuePair<Declaration, string>(efd.Declaration, efd.IdentifierName);
        }

        public Declaration Declaration => _decorated.Declaration;

        public string IdentifierName => _decorated.IdentifierName;

        public DeclarationType DeclarationType => _decorated.DeclarationType;

        public Accessibility Accessibility => _decorated.Accessibility;

        public IFieldEncapsulationAttributes EncapsulationAttributes
        {
            get => _decorated.EncapsulationAttributes;
            set => _decorated.EncapsulationAttributes = value;
        }

        public KeyValuePair<Declaration, string> TargetIDPair { get => _decorated.TargetIDPair; set => _decorated.TargetIDPair = value; }
    }

    public class EncapsulatedValueType : EncapsulateFieldDecoratorBase
    {
        private EncapsulatedValueType(IEncapsulatedFieldDeclaration efd)
            : base(efd)
        {
            EncapsulationAttributes.CanImplementLet = true;
            EncapsulationAttributes.CanImplementSet = false;
            EncapsulationAttributes.ImplementLetSetterType = true;
            EncapsulationAttributes.ImplementSetSetterType = false;
        }

        public static IEncapsulatedFieldDeclaration Decorate(IEncapsulatedFieldDeclaration efd) 
            => new EncapsulatedValueType(efd);
    }

    public class EncapsulatedUserDefinedType : EncapsulateFieldDecoratorBase
    {

        private EncapsulatedUserDefinedType(IEncapsulatedFieldDeclaration efd)
            : base(efd)
        {
            EncapsulationAttributes.CanImplementLet = true;
            EncapsulationAttributes.CanImplementSet = false;
            EncapsulationAttributes.ImplementLetSetterType = true;
            EncapsulationAttributes.ImplementSetSetterType = false;
        }

        public static IEncapsulatedFieldDeclaration Decorate(IEncapsulatedFieldDeclaration efd) 
            => new EncapsulatedUserDefinedType(efd);
    }

    public class EncapsulatedVariantType : EncapsulateFieldDecoratorBase
    {
        private EncapsulatedVariantType(IEncapsulatedFieldDeclaration efd)
            : base(efd)
        {
            EncapsulationAttributes.CanImplementLet = true;
            EncapsulationAttributes.CanImplementSet = true;
            EncapsulationAttributes.ImplementLetSetterType = true;
            EncapsulationAttributes.ImplementSetSetterType = true;
        }

        public static IEncapsulatedFieldDeclaration Decorate(IEncapsulatedFieldDeclaration efd) 
            => new EncapsulatedVariantType(efd);
    }

    public class EncapsulatedObjectType : EncapsulateFieldDecoratorBase
    {
        private EncapsulatedObjectType(IEncapsulatedFieldDeclaration efd)
            : base(efd)
        {
            EncapsulationAttributes.CanImplementLet = false;
            EncapsulationAttributes.CanImplementSet = true;
            EncapsulationAttributes.ImplementLetSetterType = false;
            EncapsulationAttributes.ImplementSetSetterType = true;
        }

        public static IEncapsulatedFieldDeclaration Decorate(IEncapsulatedFieldDeclaration efd) 
            => new EncapsulatedObjectType(efd);
    }

    public class EncapsulatedArrayType : EncapsulateFieldDecoratorBase
    {
        private EncapsulatedArrayType(IEncapsulatedFieldDeclaration efd)
            : base(efd)
        {
            EncapsulationAttributes.CanImplementLet = false;
            EncapsulationAttributes.CanImplementSet = false;
            EncapsulationAttributes.ImplementLetSetterType = false;
            EncapsulationAttributes.ImplementSetSetterType = false;
            EncapsulationAttributes.AsTypeName = Tokens.Variant;
        }

        public static IEncapsulatedFieldDeclaration Decorate(IEncapsulatedFieldDeclaration efd) 
            => new EncapsulatedArrayType(efd);
    }

    public class EncapsulatedUserDefinedTypeMember : EncapsulateFieldDecoratorBase
    {
        private IFieldEncapsulationAttributes _udtVariableAttributes;
        private bool _nameResolveProperty;
        private string _originalVariableName;
        private EncapsulatedUserDefinedTypeMember(IEncapsulatedFieldDeclaration efd, EncapsulatedUserDefinedType udtVariable, bool propertyIdentifierRequiresNameResolution)
            : base(efd)
        {
            _originalVariableName = udtVariable.IdentifierName;
            _nameResolveProperty = propertyIdentifierRequiresNameResolution;
            _udtVariableAttributes = udtVariable.EncapsulationAttributes;

            EncapsulationAttributes.NewFieldName = efd.Declaration.IdentifierName;
            EncapsulationAttributes.PropertyName = BuildPropertyName();
            EncapsulationAttributes.FieldReadWriteIdentifierFunc = ()=> $"{_udtVariableAttributes.NewFieldName}.{EncapsulationAttributes.NewFieldName}";

            efd.TargetIDPair = new KeyValuePair<Declaration, string>(efd.Declaration, $"{udtVariable.IdentifierName}.{IdentifierName}");
        }

        private string BuildPropertyName()
        {
            if (_nameResolveProperty)
            {
                var propertyPrefix = char.ToUpper(_originalVariableName[0]) + _originalVariableName.Substring(1, _originalVariableName.Length - 1);
                return $"{propertyPrefix}_{EncapsulationAttributes.FieldName}";
            }
            return EncapsulationAttributes.FieldName;
        }

        public static IEncapsulatedFieldDeclaration Decorate(IEncapsulatedFieldDeclaration efd, EncapsulatedUserDefinedType udtVariable, bool nameResolveProperty) 
            => new EncapsulatedUserDefinedTypeMember(efd, udtVariable, nameResolveProperty);
    }
}
