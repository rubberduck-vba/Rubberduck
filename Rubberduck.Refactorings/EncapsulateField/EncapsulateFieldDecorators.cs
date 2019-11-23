using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using System;
using System.Collections.Generic;
using System.Windows;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulatedFieldDeclaration
    {
        Declaration Declaration { get; }
        string IdentifierName { get; }
        string TargetID { get; set; }
        DeclarationType DeclarationType { get; }
        Accessibility Accessibility { get;}
        IFieldEncapsulationAttributes EncapsulationAttributes { set; get; }
        bool IsReadOnly { set; get; }
        bool CanBeReadWrite { set;  get; }
        string PropertyName { set; get; }
        string FieldReadWriteIdentifier { get; }
        bool EncapsulateFlag { set; get; }
        string NewFieldName { set; get; }
        string AsTypeName { get; }
        bool IsUDTMember { set; get; }
        bool HasValidEncapsulationAttributes { get; }
        Func<IEncapsulatedFieldDeclaration, bool> HasConflictsValidationFunc { set; get; }
    }

    public class EncapsulatedFieldDeclaration : IEncapsulatedFieldDeclaration
    {
        protected Declaration _decorated;
        private IFieldEncapsulationAttributes _attributes;

        public EncapsulatedFieldDeclaration(Declaration declaration)
        {
            _decorated = declaration;
            _attributes = new FieldEncapsulationAttributes(_decorated);
            TargetID = declaration.IdentifierName;
        }

        public Declaration Declaration => _decorated;

        public string IdentifierName => _decorated.IdentifierName;

        public DeclarationType DeclarationType => _decorated.DeclarationType;

        public Accessibility Accessibility => _decorated.Accessibility;

        public bool HasValidEncapsulationAttributes
        {
            get
            {
                return Declaration != null
                        && VBAIdentifierValidator.IsValidIdentifier(PropertyName, DeclarationType.Variable)
                        && !EncapsulationAttributes.PropertyName.Equals(FieldReadWriteIdentifier, StringComparison.InvariantCultureIgnoreCase)
                        && !EncapsulationAttributes.ParameterName.Equals(EncapsulationAttributes.PropertyName, StringComparison.InvariantCultureIgnoreCase);
            }
        }

        public IFieldEncapsulationAttributes EncapsulationAttributes
        {
            set => _attributes = value;
            get => _attributes;
        }

        public string TargetID { get; set; }

        public bool EncapsulateFlag
        {
            get => _attributes.EncapsulateFlag;
            set => _attributes.EncapsulateFlag = value;
        }

        public bool IsReadOnly
        {
            get => _attributes.ReadOnly;
            set => _attributes.ReadOnly = value;
        }

        public bool CanBeReadWrite { set; get; } = true;

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

        public string FieldReadWriteIdentifier
        {
            get => _attributes.FieldReadWriteIdentifier;
        }

        public string AsTypeName => _decorated.AsTypeName;

        public bool IsUDTMember { set; get; } = false;

        public Func<IEncapsulatedFieldDeclaration, bool> HasConflictsValidationFunc { set; get; }
    }

    public class EncapsulateFieldDecoratorBase : IEncapsulatedFieldDeclaration
    {

        protected IEncapsulatedFieldDeclaration _decorated;

        public EncapsulateFieldDecoratorBase(IEncapsulatedFieldDeclaration efd)
        {
            _decorated = efd;
            TargetID = efd.IdentifierName;
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
            set => _decorated.EncapsulationAttributes.NewFieldName = value;
        }

        public string FieldReadWriteIdentifier 
            => _decorated.EncapsulationAttributes.FieldReadWriteIdentifier;

        public string AsTypeName => _decorated.EncapsulationAttributes.AsTypeName;

        public bool IsUDTMember
        {
            get => _decorated.IsUDTMember;
            set => _decorated.IsUDTMember = value;
        }

        public Func<IEncapsulatedFieldDeclaration, bool> HasConflictsValidationFunc
        {
            get => _decorated.HasConflictsValidationFunc;
            set => _decorated.HasConflictsValidationFunc = value;
        }

        public bool HasValidEncapsulationAttributes
        {
            get
            {
                return _decorated.HasValidEncapsulationAttributes
                    && !(HasConflictsValidationFunc != null ? HasConflictsValidationFunc(_decorated) : false);
            }
        }
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
            _originalVariableName = udtVariable.IdentifierName;
            _nameResolveProperty = propertyIdentifierRequiresNameResolution;
            _udtVariableAttributes = udtVariable.EncapsulationAttributes;

            EncapsulationAttributes.NewFieldName = efd.Declaration.IdentifierName;
            EncapsulationAttributes.PropertyName = BuildPropertyName();
            EncapsulationAttributes.FieldReadWriteIdentifierFunc = () =>
                {
                    if (_udtVariableAttributes.EncapsulateFlag)
                    {
                        return $"{_udtVariableAttributes.NewFieldName}.{EncapsulationAttributes.NewFieldName}";
                    }
                    return $"{_udtVariableAttributes.FieldName}.{EncapsulationAttributes.NewFieldName}";
                };

            efd.TargetID = $"{udtVariable.IdentifierName}.{IdentifierName}";
            efd.IsUDTMember = true;
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
