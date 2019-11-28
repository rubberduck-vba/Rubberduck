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
        IFieldEncapsulationAttributes EncapsulationAttributes { set; get; }
        bool IsReadOnly { set; get; }
        bool CanBeReadWrite { set; get; }
        string PropertyName { set; get; }
        bool EncapsulateFlag { set; get; }
        string NewFieldName { get; }
        string AsTypeName { get; }
        bool IsUDTMember { set; get; }
        bool HasValidEncapsulationAttributes { get; }
        QualifiedModuleName QualifiedModuleName { get; }
        IEnumerable<IdentifierReference> References { get; }
    }

    public class EncapsulateFieldCandidate : IEncapsulateFieldCandidate
    {
        protected Declaration _target;
        private IFieldEncapsulationAttributes _attributes;
        private IEncapsulateFieldNamesValidator _validator;

        public EncapsulateFieldCandidate(Declaration declaration, IEncapsulateFieldNamesValidator validator)
        {
            _target = declaration;
            _attributes = new FieldEncapsulationAttributes(_target);
            _validator = validator;
            //TargetID = declaration.IdentifierName;


            var isValidAttributeSet = _validator.HasValidEncapsulationAttributes(_attributes, declaration.QualifiedModuleName, (Declaration dec) => dec.Equals(_target));
            for (var idx = 2; idx < 9 && !isValidAttributeSet; idx++)
            {
                _attributes.NewFieldName = $"{declaration.IdentifierName}{idx}";
                isValidAttributeSet = _validator.HasValidEncapsulationAttributes(_attributes, declaration.QualifiedModuleName, (Declaration dec) => dec.Equals(_target));
            }
        }

        public Declaration Declaration => _target;

        public DeclarationType DeclarationType => _target.DeclarationType;

        public bool HasValidEncapsulationAttributes 
            => _validator.HasValidEncapsulationAttributes(EncapsulationAttributes, QualifiedModuleName, (Declaration dec) => dec.Equals(_target));

        public IFieldEncapsulationAttributes EncapsulationAttributes
        {
            set => _attributes = value;
            get => _attributes;
        }

        public virtual string TargetID => _target.IdentifierName;

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

        public string AsTypeName => _target.AsTypeName;

        public bool IsUDTMember { set; get; } = false;

        public QualifiedModuleName QualifiedModuleName => Declaration.QualifiedModuleName;

        public IEnumerable<IdentifierReference> References => Declaration.References;
    }
}
