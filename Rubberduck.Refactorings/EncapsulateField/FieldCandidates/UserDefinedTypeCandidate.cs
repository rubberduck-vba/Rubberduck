using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using System;
using System.Collections.Generic;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IUserDefinedTypeCandidate : IEncapsulateFieldCandidate
    {
        IEnumerable<IUserDefinedTypeMemberCandidate> Members { get; }
        void AddMember(IUserDefinedTypeMemberCandidate member);
        bool TypeDeclarationIsPrivate { get; }
    }

    public class UserDefinedTypeCandidate : EncapsulateFieldCandidate, IUserDefinedTypeCandidate
    {
        public UserDefinedTypeCandidate(Declaration declaration)
            : base(declaration)
        {
            BackingIdentifierMutator = Declaration.AsTypeDeclaration.HasPrivateAccessibility()
                ? null
                : base.BackingIdentifierMutator;
        }

        public void AddMember(IUserDefinedTypeMemberCandidate member)
        {
            _udtMembers.Add(member);
        }

        private List<IUserDefinedTypeMemberCandidate> _udtMembers = new List<IUserDefinedTypeMemberCandidate>();
        public IEnumerable<IUserDefinedTypeMemberCandidate> Members => _udtMembers;

        public bool TypeDeclarationIsPrivate
            => Declaration.AsTypeDeclaration?.HasPrivateAccessibility() ?? false;

        public override string BackingIdentifier =>
            BackingIdentifierMutator is null
                ? _fieldAndProperty.TargetFieldName
                : _fieldAndProperty.Field;

        public override Action<string> BackingIdentifierMutator { get; } 

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
                base.EncapsulateFlag = value;
                if (TypeDeclarationIsPrivate)
                {
                    foreach (var member in Members)
                    {
                        member.EncapsulateFlag = value;
                    }
                }
            }
            get => base.EncapsulateFlag;
        }

        public override bool Equals(object obj)
            => (obj is IUserDefinedTypeCandidate udt && udt.TargetID.Equals(TargetID));

        public override int GetHashCode() => base.GetHashCode();
    }
}
