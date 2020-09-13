using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldAsUDTMemberCandidate : IEncapsulateFieldCandidate
    {
        IObjectStateUDT ObjectStateUDT { set; get; }
        IEncapsulateFieldCandidate WrappedCandidate { get; }
    }

    public class EncapsulateFieldAsUDTMemberCandidate : IEncapsulateFieldAsUDTMemberCandidate
    {
        private int _hashCode;
        private readonly string _uniqueID;
        private IEncapsulateFieldCandidate _wrapped;
        public EncapsulateFieldAsUDTMemberCandidate(IEncapsulateFieldCandidate candidate, IObjectStateUDT objStateUDT)
        {
            _wrapped = candidate;
            PropertyIdentifier = _wrapped.PropertyIdentifier;
            ObjectStateUDT = objStateUDT;
            _uniqueID = BuildUniqueID(candidate, objStateUDT);
            _hashCode = _uniqueID.GetHashCode();
        }

        public IEncapsulateFieldCandidate WrappedCandidate => _wrapped;

        private IObjectStateUDT _objectStateUDT;
        public IObjectStateUDT ObjectStateUDT
        {
            set
            {
                _objectStateUDT = value;
                if (_objectStateUDT?.Declaration == Declaration)
                {
                    EncapsulateFlag = false;
                }
            }
            get => _objectStateUDT;
        }

        public string TargetID => _wrapped.TargetID;

        public Declaration Declaration => _wrapped.Declaration;

        public bool EncapsulateFlag
        {
            set => _wrapped.EncapsulateFlag = value;
            get => _wrapped.EncapsulateFlag;
        }

        public string PropertyIdentifier
        {
            set => _wrapped.PropertyIdentifier = value;
            get => _wrapped.PropertyIdentifier;
        }

        public string PropertyAsTypeName => _wrapped.PropertyAsTypeName;

        public string BackingIdentifier
        {
            set { }
            get => PropertyIdentifier;
        }
        public string BackingAsTypeName => Declaration.AsTypeName;

        public bool CanBeReadWrite
        {
            set => _wrapped.CanBeReadWrite = value;
            get => _wrapped.CanBeReadWrite;
        }

        public bool ImplementLet => _wrapped.ImplementLet;

        public bool ImplementSet => _wrapped.ImplementSet;

        public bool IsReadOnly
        {
            set => _wrapped.IsReadOnly = value;
            get => _wrapped.IsReadOnly;
        }

        public IEncapsulateFieldConflictFinder ConflictFinder
        {
            set => _wrapped.ConflictFinder = value;
            get => _wrapped.ConflictFinder;
        }

        public string IdentifierName => _wrapped.IdentifierName;

        public QualifiedModuleName QualifiedModuleName => _wrapped.QualifiedModuleName;

        public string AsTypeName => _wrapped.AsTypeName;

        public bool TryValidateEncapsulationAttributes(out string errorMessage)
        {
            errorMessage = string.Empty;
            if (!_wrapped.EncapsulateFlag)
            {
                return true;
            }

            if (_wrapped is IArrayCandidate ac)
            {
                if (ac.HasExternalRedimOperation(out errorMessage))
                {
                    return false;
                }
            }
            return ConflictFinder.TryValidateEncapsulationAttributes(this, out errorMessage);
        }

        public override bool Equals(object obj)
        {
            return obj != null
                && obj is EncapsulateFieldAsUDTMemberCandidate convertWrapper
                && BuildUniqueID(convertWrapper, convertWrapper.ObjectStateUDT) == _uniqueID;
        }

        public override int GetHashCode() => _hashCode;

        private static string BuildUniqueID(IEncapsulateFieldCandidate candidate, IObjectStateUDT field) 
            => $"{candidate.QualifiedModuleName.Name}.{field.IdentifierName}.{candidate.IdentifierName}";
    }
}
