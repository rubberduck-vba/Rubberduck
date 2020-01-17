using Rubberduck.Refactorings.EncapsulateField;

namespace Rubberduck.UI.Refactorings.EncapsulateField
{
    public interface IEncapsulatedFieldViewData
    {
        string TargetID { get; }
        string PropertyName { set; get; }
        bool EncapsulateFlag { set; get; }
        bool IsReadOnly { set; get; }
        bool CanBeReadWrite { get; }
        bool HasValidEncapsulationAttributes { get; }
        string TargetDeclarationExpression { set; get; }
        bool IsPrivateUserDefinedType { get; }
        bool IsRequiredToBeReadOnly { get; }
        string ValidationErrorMessage { get; }
        bool TryValidateEncapsulationAttributes(out string errorMessage);
    }

    public class ViewableEncapsulatedField : IEncapsulatedFieldViewData
    {
        private IEncapsulateFieldCandidate _efd;
        private readonly int _hashCode;
        public ViewableEncapsulatedField(IEncapsulateFieldCandidate efd)
        {
            _efd = efd;
            _hashCode = efd.TargetID.GetHashCode();
        }

        public override bool Equals(object obj)
        {
            if (obj is IEncapsulatedFieldViewData vd)
            {
                return vd.TargetID.Equals(TargetID);
            }
            return false;
        }

        public override int GetHashCode()
        {
            return _hashCode;
        }

        public bool HasValidEncapsulationAttributes
        {
            get
            {
                if (!TryValidateEncapsulationAttributes(out var errorMessage))
                {
                    _errorMessage = errorMessage;
                    return false;
                }
                return true;
            }
        }

        public bool TryValidateEncapsulationAttributes(out string errorMessage)
        {
            errorMessage = string.Empty;
            return _efd.TryValidateEncapsulationAttributes(out errorMessage);
        }

        private string _errorMessage;
        public string ValidationErrorMessage => _errorMessage;

        public string TargetID { get => _efd.TargetID; }

        public bool IsReadOnly { get => _efd.IsReadOnly; set => _efd.IsReadOnly = value; }

        public bool CanBeReadWrite => _efd.CanBeReadWrite;

        public string PropertyName { get => _efd.PropertyIdentifier; set => _efd.PropertyIdentifier = value; }

        public bool EncapsulateFlag { get => _efd.EncapsulateFlag; set => _efd.EncapsulateFlag = value; }

        public bool IsPrivateUserDefinedType => _efd is IUserDefinedTypeCandidate udt && udt.TypeDeclarationIsPrivate;

        public bool IsRequiredToBeReadOnly => !_efd.CanBeReadWrite;

        private string _targetDeclarationExpressions;
        public string TargetDeclarationExpression
        {
            set => _targetDeclarationExpressions = value;
            get => $"{_efd.Declaration.Accessibility} {_efd.Declaration.Context.GetText()}";
        }
    }
}
