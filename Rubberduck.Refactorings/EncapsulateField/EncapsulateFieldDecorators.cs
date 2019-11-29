using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulatedUserDefinedTypeField : EncapsulateFieldCandidate
    {
        public List<IEncapsulateFieldCandidate> Members { set; get; }
        public EncapsulatedUserDefinedTypeField(Declaration declaration, IEncapsulateFieldNamesValidator validator)
            : base(declaration, validator)
        {
        }
    }

    public class EncapsulatedUserDefinedTypeMember : IEncapsulateFieldCandidate //: EncapsulateFieldDecoratorBase
    {
        private IEncapsulateFieldCandidate _candidate;
        public EncapsulatedUserDefinedTypeMember(IEncapsulateFieldCandidate candidate, IEncapsulateFieldCandidate udtVariable, bool propertyNameRequiresParentIdentiier)
        {
            _candidate = candidate;
            Parent = udtVariable;
            NameResolveProperty = propertyNameRequiresParentIdentiier;

            EncapsulationAttributes.PropertyName = NameResolveProperty 
                ? $"{Parent.IdentifierName.Capitalize()}_{IdentifierName}" 
                : IdentifierName;

            //var fea = EncapsulationAttributes as FieldEncapsulationAttributes;
            (EncapsulationAttributes as FieldEncapsulationAttributes).FieldReferenceExpressionFunc = () => AsWithMemberExpression;
            candidate.IsUDTMember = true;
        }

        public string AsWithMemberExpression
        {
            get
            {
                var prefix = Parent.EncapsulateFlag
                               ? Parent.NewFieldName
                               : Parent.IdentifierName;

                return $"{prefix}.{NewFieldName}";
            }
        }

        public IEncapsulateFieldCandidate Parent { private set; get; }

        public bool NameResolveProperty { set; get; }

        public bool IsExistingDeclaration => _candidate.IsExistingDeclaration;

        public Declaration Declaration => _candidate.Declaration;

        public DeclarationType DeclarationType => _candidate.DeclarationType;

        public IFieldEncapsulationAttributes EncapsulationAttributes
        {
            get => _candidate.EncapsulationAttributes;
            set => _candidate.EncapsulationAttributes = value;
        }

        public bool IsReadOnly
        {
            get => _candidate.EncapsulationAttributes.ReadOnly;
            set => _candidate.EncapsulationAttributes.ReadOnly = value;
        }

        public bool EncapsulateFlag
        {
            get => _candidate.EncapsulateFlag;
            set => _candidate.EncapsulateFlag = value;
        }

        public bool CanBeReadWrite
        {
            get => _candidate.CanBeReadWrite;
            set => _candidate.CanBeReadWrite = value;
        }

        public string PropertyName
        {
            get => _candidate.EncapsulationAttributes.PropertyName;
            set => _candidate.EncapsulationAttributes.PropertyName = value;
        }

        public string NewFieldName
        {
            get => _candidate.EncapsulationAttributes.NewFieldName;
        }

        public string AsTypeName => _candidate.EncapsulationAttributes.AsTypeName;

        public bool IsUDTMember
        {
            get => _candidate.IsUDTMember;
            set => _candidate.IsUDTMember = value;
        }

        public bool HasValidEncapsulationAttributes
            => _candidate.HasValidEncapsulationAttributes;

        public QualifiedModuleName QualifiedModuleName => _candidate.QualifiedModuleName;

        public IEnumerable<IdentifierReference> References => _candidate.References;

        public string TargetID => $"{Parent.IdentifierName}.{IdentifierName}";

        public string IdentifierName => _candidate.IdentifierName;

        public string ParameterName => _candidate.ParameterName;// throw new NotImplementedException();

        public string FieldReferenceExpression => _candidate.FieldReferenceExpression; // throw new NotImplementedException();

        public bool ImplementLetSetterType { get => _candidate.ImplementLetSetterType; /*throw new NotImplementedException();*/ set => _candidate.ImplementLetSetterType = value; } // throw new NotImplementedException(); }
        public bool ImplementSetSetterType { get => _candidate.ImplementSetSetterType; /*throw new NotImplementedException();*/ set => _candidate.ImplementSetSetterType = value; } // throw new NotImplementedException(); }

        //private string BuildPropertyName() => NameResolveProperty ? $"{Parent.IdentifierName.Capitalize()}_{IdentifierName}" : IdentifierName;
        //{
        //    if (NameResolveProperty)
        //    {
        //        return $"{Parent.IdentifierName.Capitalize()}_{IdentifierName}";
        //    }
        //    return IdentifierName;
        //}
    }
}
