using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulatedFieldDeclaration
    {
        Declaration Declaration { get; }
        string TargetID { get; set; }
        IFieldEncapsulationAttributes EncapsulationAttributes { set; get; }
        bool IsReadOnly { set; get; }
        bool CanBeReadWrite { set; get; }
        string PropertyName { set; get; }
        string FieldReadWriteIdentifier { get; }
        bool EncapsulateFlag { set; get; }
        string NewFieldName { get; }
        string AsTypeName { get; }
        bool IsUDTMember { set; get; }
        bool HasValidEncapsulationAttributes { get; }
    }

    public class EncapsulatedFieldDeclaration : IEncapsulatedFieldDeclaration
    {
        protected Declaration _decorated;
        private IFieldEncapsulationAttributes _attributes;
        private IEncapsulateFieldNamesValidator _validator;

        public EncapsulatedFieldDeclaration(Declaration declaration, IEncapsulateFieldNamesValidator validator)
        {
            _decorated = declaration;
            _attributes = new FieldEncapsulationAttributes(_decorated);
            _validator = validator;
            TargetID = declaration.IdentifierName;
            SetNonConflictingEncapsulationAttributes();
        }

        private void SetNonConflictingEncapsulationAttributes()
        {
            var isValid = _validator.HasValidEncapsulationAttributes(this);
            if (!isValid)
            {
                var attributes = new ClientEncapsulationAttributes(Declaration.IdentifierName);
                if (IsConflictingAttributes(attributes))
                {
                    var hasConflict = true;
                    for (var idx = 2; idx < 9 && hasConflict; idx++)
                    {
                        attributes = new ClientEncapsulationAttributes($"{Declaration.IdentifierName}{idx}");
                        hasConflict = IsConflictingAttributes(attributes);
                    }
                    PropertyName = attributes.PropertyName;
                }
            }
        }

        private bool IsConflictingAttributes(ClientEncapsulationAttributes attributes)
        {
            var isConflictingFieldName = _validator.IsConflictingMemberName(attributes.NewFieldName, Declaration.QualifiedModuleName, Declaration.DeclarationType);
            var isConflictingPropertyName = _validator.IsConflictingMemberName(attributes.PropertyName, Declaration.QualifiedModuleName, DeclarationType.Member);
            return isConflictingFieldName || isConflictingPropertyName;
        }

        public Declaration Declaration => _decorated;

        public bool HasValidEncapsulationAttributes
        {
            get
            {
                return _validator.HasValidEncapsulationAttributes(this);
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
        }

        public string FieldReadWriteIdentifier
        {
            get => _attributes.FieldReadWriteIdentifier;
        }

        public string AsTypeName => _decorated.AsTypeName;

        public bool IsUDTMember { set; get; } = false;
    }
}
