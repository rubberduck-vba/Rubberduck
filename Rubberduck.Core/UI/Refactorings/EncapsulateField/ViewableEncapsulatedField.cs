using Rubberduck.Refactorings.EncapsulateField;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Rubberduck.UI.Refactorings.EncapsulateField
{
    public interface IEncapsulatedFieldViewData
    {
        string TargetID { get; set; }
        string PropertyName { set; get; }
        string NewFieldName { set; get; }
        bool EncapsulateFlag { set; get; }
        bool IsReadOnly { set; get; }
        bool CanBeReadWrite { get; }
        bool IsEditableReadWriteFieldIdentifier { get; }
        Visibility FieldNameVisibility { get; }
        Visibility PropertyNameVisibility { get; }
        bool HasValidEncapsulationAttributes { get; }
        string AsTypeName { get; }
    }

    public class ViewableEncapsulatedField : IEncapsulatedFieldViewData
    {
        private IEncapsulatedFieldDeclaration _efd;
        public ViewableEncapsulatedField(IEncapsulatedFieldDeclaration efd)
        {
            _efd = efd;
        }

        public Visibility FieldNameVisibility => _efd.IsUDTMember || !_efd.EncapsulateFlag ? Visibility.Collapsed : Visibility.Visible;
        public Visibility PropertyNameVisibility => !_efd.EncapsulateFlag ? Visibility.Collapsed : Visibility.Visible;
        public bool HasValidEncapsulationAttributes => _efd.HasValidEncapsulationAttributes;
        public string TargetID { get => _efd.TargetID; set => _efd.TargetID = value; }
        public bool IsReadOnly { get => _efd.IsReadOnly; set => _efd.IsReadOnly = value; }
        public bool CanBeReadWrite => _efd.CanBeReadWrite;
        public string PropertyName { get => _efd.PropertyName; set => _efd.PropertyName = value; }
        public bool IsEditableReadWriteFieldIdentifier { get => !_efd.IsUDTMember; } // set => _efd.IsEditableReadWriteFieldIdentifier = value; }
        public bool EncapsulateFlag { get => _efd.EncapsulateFlag; set => _efd.EncapsulateFlag = value; }
        public string NewFieldName { get => _efd.NewFieldName; set => _efd.NewFieldName = value; }
        //TODO: Change name of AsTypeName property to FieldDescriptor(?)  -> and does it belong on IEncapsulatedField?
        public string AsTypeName
        {
            //(Variable: Integer Array)
            //(Variable: Long)
            //(UserDefinedType Member: Long)
            get
            {
                var prefix = string.Empty;

                var descriptor = string.Empty;
                if (_efd.IsUDTMember)
                {
                    prefix = "UserDefinedType";
                }
                else
                {
                    prefix = "Variable";
                }

                descriptor = $"{prefix}: {_efd.Declaration.AsTypeName}";
                if (_efd.Declaration.IsArray)
                {
                    descriptor = $"{descriptor} Array";
                }
                return descriptor;
            }
        }
    }
}
