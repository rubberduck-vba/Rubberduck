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
        string IdentifierName { get; }
        IFieldEncapsulationAttributes EncapsulationAttributes { set; get; }
        bool IsReadOnly { set; get; }
        bool CanBeReadWrite { set; get; }
        string PropertyName { set; get; }
        bool EncapsulateFlag { set; get; }
        string NewFieldName { get; }
        string AsTypeName { get; }
        string ParameterName { get; }
        string FieldReferenceExpression { get; }
        bool IsUDTMember { set; get; }
        bool HasValidEncapsulationAttributes { get; }
        QualifiedModuleName QualifiedModuleName { get; }
        IEnumerable<IdentifierReference> References { get; }
        bool ImplementLetSetterType { get; set; }
        bool ImplementSetSetterType { get; set; }
        bool IsExistingDeclaration { get; }
    }


    ////string TargetName { get; }
    ////string PropertyName { get; set; }
    ////bool ReadOnly { get; set; }
    ////bool EncapsulateFlag { get; set; }
    ////string NewFieldName { set; get; }
    //string FieldReferenceExpression { get; }
    ////string AsTypeName { get; set; }
    //string ParameterName { get; }
    //bool ImplementLetSetterType { get; set; }
    //bool ImplementSetSetterType { get; set; }
    //bool FieldNameIsExemptFromValidation { get; }

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


            var isValidAttributeSet = _validator.HasValidEncapsulationAttributes(_attributes, declaration.QualifiedModuleName, new Declaration[] { _target }); // (Declaration dec) => dec.Equals(_target));
            for (var idx = 2; idx < 9 && !isValidAttributeSet; idx++)
            {
                _attributes.NewFieldName = $"{declaration.IdentifierName}{idx}";
                isValidAttributeSet = _validator.HasValidEncapsulationAttributes(_attributes, declaration.QualifiedModuleName, new Declaration[] { _target }); //(Declaration dec) => dec.Equals(_target));
            }
        }

        public EncapsulateFieldCandidate(IFieldEncapsulationAttributes attributes, IEncapsulateFieldNamesValidator validator)
        {
            _target = null;
            _attributes = attributes; // new FieldEncapsulationAttributes(identifier, asTypeName);
            _validator = validator;
        }

        public Declaration Declaration => _target;

        public bool IsExistingDeclaration => _target != null;

        public DeclarationType DeclarationType => _target?.DeclarationType ?? DeclarationType.Variable;

        public bool HasValidEncapsulationAttributes
        {
            get
            {
                var ignore = _target != null ? new Declaration[] { _target } : Enumerable.Empty<Declaration>();
                return _validator.HasValidEncapsulationAttributes(EncapsulationAttributes, QualifiedModuleName, ignore); //(Declaration dec) => dec.Equals(_target));
            }
       }

    public IFieldEncapsulationAttributes EncapsulationAttributes
        {
            set => _attributes = value;
            get => _attributes;
        }

        public virtual string TargetID => _target?.IdentifierName ?? _attributes.Identifier;

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

        public string AsTypeName => _target?.AsTypeName ?? _attributes.AsTypeName;

        //TODO: This needs to be readonly
        public bool IsUDTMember { set; get; } = false;

        public QualifiedModuleName QualifiedModuleName => Declaration?.QualifiedModuleName ?? _attributes.QualifiedModuleName;

        public IEnumerable<IdentifierReference> References => Declaration?.References ?? Enumerable.Empty<IdentifierReference>();

        public string IdentifierName => Declaration?.IdentifierName ?? _attributes.Identifier;

        public string ParameterName => _attributes.ParameterName;// throw new NotImplementedException();

        public string FieldReferenceExpression => _attributes.FieldReferenceExpression; // throw new NotImplementedException();

        public bool ImplementLetSetterType { get => _attributes.ImplementLetSetterType; /*throw new NotImplementedException();*/ set => _attributes.ImplementLetSetterType = value; } // throw new NotImplementedException(); }
        public bool ImplementSetSetterType { get => _attributes.ImplementSetSetterType; /*throw new NotImplementedException();*/ set => _attributes.ImplementSetSetterType = value; } // throw new NotImplementedException(); }
    }
}
