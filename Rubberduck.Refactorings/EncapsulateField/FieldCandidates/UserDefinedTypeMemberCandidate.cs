using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IUserDefinedTypeMemberCandidate : IEncapsulateFieldCandidate
    {
        IUserDefinedTypeCandidate UDTField { get; }
        IEncapsulateFieldCandidate WrappedCandidate { get; }
    }

    public class UserDefinedTypeMemberCandidate : IUserDefinedTypeMemberCandidate
    {
        private readonly int _hashCode;

        public UserDefinedTypeMemberCandidate(IEncapsulateFieldCandidate candidate, IUserDefinedTypeCandidate udtField)
        {
            WrappedCandidate = candidate;
            UDTField = udtField;
            PropertyIdentifier = IdentifierName;
            BackingIdentifier = IdentifierName;
            _hashCode = TargetID.GetHashCode();
        }

        public IEncapsulateFieldCandidate WrappedCandidate { private set; get; }

        public string AsTypeName => WrappedCandidate.AsTypeName;

        public IUserDefinedTypeCandidate UDTField { private set; get; }

        public IEncapsulateFieldConflictFinder ConflictFinder
        {
            set => WrappedCandidate.ConflictFinder = value;
            get => WrappedCandidate.ConflictFinder;
        }

        public string TargetID => $"{UDTField.IdentifierName}.{IdentifierName}";

        public string IdentifierForReference(IdentifierReference idRef)
            => PropertyIdentifier;

        public string PropertyIdentifier { set; get; }

        public string BackingIdentifier { get; }

        public Action<string> BackingIdentifierMutator { get; } = null;

        public Declaration Declaration => WrappedCandidate.Declaration;

        public string IdentifierName => WrappedCandidate.IdentifierName;

        public bool TryValidateEncapsulationAttributes(out string errorMessage)
        {
            errorMessage = string.Empty;
            return true;
        }

        public bool IsReadOnly
        {
            set => WrappedCandidate.IsReadOnly = value;
            get => WrappedCandidate.IsReadOnly;
        }

        public bool IsAssignedExternally => WrappedCandidate.IsAssignedExternally;

        private bool _encapsulateFlag;
        public bool EncapsulateFlag
        {
            set
            {
                if (WrappedCandidate is IUserDefinedTypeCandidate udt && udt.TypeDeclarationIsPrivate)
                {
                    foreach (var member in udt.Members)
                    {
                        member.EncapsulateFlag = value;
                    }
                    return;
                }

                var valueChanged = _encapsulateFlag != value;
                _encapsulateFlag = value;

                PropertyIdentifier = WrappedCandidate.PropertyIdentifier;

                if (_encapsulateFlag && valueChanged && ConflictFinder != null)
                {
                    ConflictFinder.AssignNoConflictIdentifiers(this);
                }

                if (!_encapsulateFlag)
                {
                    WrappedCandidate.EncapsulateFlag = value;
                }

            }
            get => _encapsulateFlag;
        }

        public bool CanBeReadWrite => !Declaration.IsArray;

        public bool HasValidEncapsulationAttributes => true;

        public QualifiedModuleName QualifiedModuleName
            => WrappedCandidate.QualifiedModuleName;

        public string PropertyAsTypeName => WrappedCandidate.PropertyAsTypeName;

        public override bool Equals(object obj)
        {
            return obj != null
                && obj is IUserDefinedTypeMemberCandidate udtMember
                && udtMember.QualifiedModuleName == QualifiedModuleName
                && udtMember.TargetID == TargetID;
        }

        public override int GetHashCode() => _hashCode;
    }
}
