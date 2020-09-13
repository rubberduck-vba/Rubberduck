using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using System.Collections.Generic;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IUserDefinedTypeCandidate : IEncapsulateFieldCandidate
    {
        IEnumerable<IUserDefinedTypeMemberCandidate> Members { get; }
        void AddMember(IUserDefinedTypeMemberCandidate member);
        bool TypeDeclarationIsPrivate { get; }
        bool IsObjectStateUDTCandidate { set; get; }
        bool IsSelectedObjectStateUDT { set; get; }
    }

    public class UserDefinedTypeCandidate : EncapsulateFieldCandidate, IUserDefinedTypeCandidate
    {
        public UserDefinedTypeCandidate(Declaration declaration)
            : base(declaration)
        {}

        public void AddMember(IUserDefinedTypeMemberCandidate member)
        {
            _udtMembers.Add(member);
        }

        private List<IUserDefinedTypeMemberCandidate> _udtMembers = new List<IUserDefinedTypeMemberCandidate>();
        public IEnumerable<IUserDefinedTypeMemberCandidate> Members => _udtMembers;

        public bool TypeDeclarationIsPrivate
            => Declaration.AsTypeDeclaration?.HasPrivateAccessibility() ?? false;

        public bool IsSelectedObjectStateUDT { set; get; }

        public bool IsObjectStateUDTCandidate { set; get; }

        public override string BackingIdentifier
        {
            get => TypeDeclarationIsPrivate ? _fieldAndProperty.TargetFieldName : _fieldAndProperty.Field;
            set => _fieldAndProperty.Field = value;
        }

        private IEncapsulateFieldConflictFinder _conflictsFinder;
        public override IEncapsulateFieldConflictFinder ConflictFinder
        {
            set
            {
                _conflictsFinder = value;
                foreach (var member in Members)
                {
                    member.ConflictFinder = value;
                }
            }
            get => _conflictsFinder;
        }

        private bool _isReadOnly;
        public override bool IsReadOnly
        {
            get => _isReadOnly;
            set
            {
                _isReadOnly = value;
                foreach ( var member in Members)
                {
                    member.IsReadOnly = value;
                }
            }
        }

        public override bool EncapsulateFlag
        {
            set
            {
                if (TypeDeclarationIsPrivate)
                {
                    foreach (var member in Members)
                    {
                        member.EncapsulateFlag = value;
                    }
                }
                base.EncapsulateFlag = value;
            }
            get => base.EncapsulateFlag;
        }

        public override bool Equals(object obj)
        {
            if (obj is IUserDefinedTypeCandidate udt)
            {
                return udt.TargetID.Equals(TargetID);
            }
            return false;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }
}
