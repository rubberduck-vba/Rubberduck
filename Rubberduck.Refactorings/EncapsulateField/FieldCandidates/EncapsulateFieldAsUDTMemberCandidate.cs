using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldAsUDTMemberCandidate : IEncapsulateFieldCandidate
    {
        IObjectStateUDT ObjectStateUDT { set; get; }
        IEncapsulateFieldCandidate WrappedCandidate { get; }
        string UserDefinedTypeMemberIdentifier { set; get; }
    }

    /// <summary>
    /// EncapsulateFieldAsUDTMemberCandidate wraps an IEncapusulateFieldCandidate instance
    /// for the purposes of declaring it as a new UserDefinedTypeMember
    /// within an existing or new UserDefinedType
    /// </summary>
    public class EncapsulateFieldAsUDTMemberCandidate : IEncapsulateFieldAsUDTMemberCandidate
    {
        private readonly int _hashCode;
        private IEncapsulateFieldCandidate _wrapped;
        public EncapsulateFieldAsUDTMemberCandidate(IEncapsulateFieldCandidate candidate, IObjectStateUDT objStateUDT)
        {
            _wrapped = candidate;
            ObjectStateUDT = objStateUDT;
            _hashCode = $"{candidate.QualifiedModuleName.Name}.{candidate.IdentifierName}".GetHashCode();
        }

        public IEncapsulateFieldCandidate WrappedCandidate => _wrapped;

        private IObjectStateUDT _objectStateUDT;
        public IObjectStateUDT ObjectStateUDT
        {
            set
            {
                _objectStateUDT = value;
                if (_objectStateUDT?.Declaration == _wrapped.Declaration)
                {
                    //Cannot wrap itself if it is used as the ObjectStateUDT 
                    _wrapped.EncapsulateFlag = false;
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

        public string UserDefinedTypeMemberIdentifier
        {
            set => PropertyIdentifier = value;
            get => PropertyIdentifier;
        }

        public string PropertyIdentifier
        {
            set => _wrapped.PropertyIdentifier = value;
            get => _wrapped.PropertyIdentifier;
        }

        public virtual Action<string> BackingIdentifierMutator { get; } = null;

        public string BackingIdentifier => PropertyIdentifier;

        public string PropertyAsTypeName => _wrapped.PropertyAsTypeName;

        public bool CanBeReadWrite => _wrapped.CanBeReadWrite;

        public bool IsReadOnly
        {
            set => _wrapped.IsReadOnly = value;
            get => _wrapped.IsReadOnly;
        }

        public bool IsAssignedExternally => _wrapped.IsAssignedExternally;

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
            (bool IsValid, string ErrorMsg) = ConflictFinder?.ValidateEncapsulationAttributes(this) ?? (true, string.Empty);
            errorMessage = ErrorMsg;
            return IsValid;
        }

        public override bool Equals(object obj)
        {
            return obj != null
                && obj is EncapsulateFieldAsUDTMemberCandidate convertWrapper
                && convertWrapper.QualifiedModuleName == QualifiedModuleName
                && convertWrapper.IdentifierName == IdentifierName;
        }

        public override int GetHashCode() => _hashCode;
    }
}
