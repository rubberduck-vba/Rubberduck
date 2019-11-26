using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using System;
using System.Collections.Generic;
using System.Windows;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldDecoratorBase : IEncapsulatedFieldDeclaration
    {
        protected IEncapsulatedFieldDeclaration _decorated;

        public EncapsulateFieldDecoratorBase(IEncapsulatedFieldDeclaration efd)
        {
            _decorated = efd;
            TargetID = efd.Declaration.IdentifierName;
        }

        public Declaration Declaration => _decorated.Declaration;

        public IFieldEncapsulationAttributes EncapsulationAttributes
        {
            get => _decorated.EncapsulationAttributes;
            set => _decorated.EncapsulationAttributes = value;
        }

        public string TargetID
        {
            get => _decorated.TargetID;
            set => _decorated.TargetID = value;
        }

        public bool IsReadOnly
        {
            get => _decorated.EncapsulationAttributes.ReadOnly;
            set => _decorated.EncapsulationAttributes.ReadOnly = value;
        }

        public bool EncapsulateFlag
        {
            get => _decorated.EncapsulateFlag;
            set => _decorated.EncapsulateFlag = value;
        }

        public bool CanBeReadWrite
        {
            get => _decorated.CanBeReadWrite;
            set => _decorated.CanBeReadWrite = value;
        }

        public string PropertyName
        {
            get => _decorated.EncapsulationAttributes.PropertyName;
            set => _decorated.EncapsulationAttributes.PropertyName = value;
        }

        public string NewFieldName
        {
            get => _decorated.EncapsulationAttributes.NewFieldName;
        }

        public string AsTypeName => _decorated.EncapsulationAttributes.AsTypeName;

        public bool IsUDTMember
        {
            get => _decorated.IsUDTMember;
            set => _decorated.IsUDTMember = value;
        }

        public bool HasValidEncapsulationAttributes 
            => _decorated.HasValidEncapsulationAttributes;
    }

    public class EncapsulatedValueType : EncapsulateFieldDecoratorBase
    {
        private EncapsulatedValueType(IEncapsulatedFieldDeclaration efd)
            : base(efd)
        {
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
            EncapsulationAttributes.ImplementLetSetterType = false;
            EncapsulationAttributes.ImplementSetSetterType = false;
            EncapsulationAttributes.AsTypeName = Tokens.Variant;
            CanBeReadWrite = false;
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
            _originalVariableName = udtVariable.Declaration.IdentifierName;
            _nameResolveProperty = propertyIdentifierRequiresNameResolution;
            _udtVariableAttributes = udtVariable.EncapsulationAttributes;

            EncapsulationAttributes.PropertyName = BuildPropertyName();
            if (EncapsulationAttributes is FieldEncapsulationAttributes fea)
            {
                fea.FieldReferenceExpressionFunc =
                 () =>  { var prefix = _udtVariableAttributes.EncapsulateFlag
                                         ? _udtVariableAttributes.NewFieldName
                                         : _udtVariableAttributes.TargetName;

                            return $"{prefix}.{EncapsulationAttributes.NewFieldName}";
                        };
            }

            efd.TargetID = $"{udtVariable.Declaration.IdentifierName}.{Declaration.IdentifierName}";
            efd.IsUDTMember = true;
        }

        public static IEncapsulatedFieldDeclaration Decorate(IEncapsulatedFieldDeclaration efd, EncapsulatedUserDefinedType udtVariable, bool nameResolveProperty) 
            => new EncapsulatedUserDefinedTypeMember(efd, udtVariable, nameResolveProperty);

        private string BuildPropertyName()
        {
            if (_nameResolveProperty)
            {
                var propertyPrefix = char.ToUpper(_originalVariableName[0]) + _originalVariableName.Substring(1, _originalVariableName.Length - 1);
                return $"{propertyPrefix}_{EncapsulationAttributes.TargetName}";
            }
            return EncapsulationAttributes.TargetName;
        }
    }
}
