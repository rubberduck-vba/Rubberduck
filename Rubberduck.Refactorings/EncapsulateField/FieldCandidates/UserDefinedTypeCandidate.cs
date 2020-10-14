using Rubberduck.Parsing.Grammar;
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
        bool TypeDeclarationIsPrivate { set; get; }
        bool CanBeObjectStateUDT { set; get; }
        bool IsSelectedObjectStateUDT { set; get; }
    }

    public class UserDefinedTypeCandidate : EncapsulateFieldCandidate, IUserDefinedTypeCandidate
    {
        public UserDefinedTypeCandidate(Declaration declaration, IValidateVBAIdentifiers identifierValidator)
            : base(declaration, identifierValidator)
        {
        }

        public void AddMember(IUserDefinedTypeMemberCandidate member)
        {
            _udtMembers.Add(member);
        }

        private List<IUserDefinedTypeMemberCandidate> _udtMembers = new List<IUserDefinedTypeMemberCandidate>();
        public IEnumerable<IUserDefinedTypeMemberCandidate> Members => _udtMembers;

        private bool _isPrivate;
        public bool TypeDeclarationIsPrivate
        {
            set => _isPrivate = value;
            get => Declaration.AsTypeDeclaration?.HasPrivateAccessibility() ?? false;
        }

        public bool IsSelectedObjectStateUDT { set; get; }

        private bool _canBeObjectStateUDT;
        public bool CanBeObjectStateUDT
        {
            set => _canBeObjectStateUDT = value;
            get => _canBeObjectStateUDT;
        }

        public override string BackingIdentifier
        {
            get => TypeDeclarationIsPrivate ? _fieldAndProperty.TargetFieldName : _fieldAndProperty.Field;
            set => _fieldAndProperty.Field = value;
        }

        private IValidateVBAIdentifiers _namesValidator;
        public override IValidateVBAIdentifiers NameValidator
        {
            set
            {
                _namesValidator = value;
                foreach (var member in Members)
                {
                    member.NameValidator = value;
                }
            }
            get => _namesValidator;
        }

        private IEncapsulateFieldConflictFinder _conflictsValidator;
        public override IEncapsulateFieldConflictFinder ConflictFinder
        {
            set
            {
                _conflictsValidator = value;
                foreach (var member in Members)
                {
                    member.ConflictFinder = value;
                }
            }
            get => _conflictsValidator;
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

        protected override string IdentifierForLocalReferences(IdentifierReference idRef)
        {
            if (idRef.Context.Parent.Parent is VBAParser.WithStmtContext wsc)
            {
                return BackingIdentifier;
            }

            return TypeDeclarationIsPrivate ? BackingIdentifier : PropertyIdentifier;
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

        public override IEnumerable<PropertyAttributeSet> PropertyAttributeSets
        {
            get
            {
                if (TypeDeclarationIsPrivate)
                {
                    var specs = new List<PropertyAttributeSet>();
                    foreach (var member in Members)
                    {
                        var sets = member.PropertyAttributeSets;
                        var modifiedSets = new List<PropertyAttributeSet>();
                        PropertyAttributeSet newSet;
                        foreach (var set in sets)
                        {
                            newSet = set;
                            newSet.BackingField = $"{BackingIdentifier}.{set.BackingField}";
                            modifiedSets.Add(newSet);
                        }
                        specs.AddRange(modifiedSets);
                    }
                    return specs;
                }
                return new List<PropertyAttributeSet>() { AsPropertyAttributeSet };
            }
        }

        protected override PropertyAttributeSet AsPropertyAttributeSet
        {
            get
            {
                return new PropertyAttributeSet()
                {
                    PropertyName = PropertyIdentifier,
                    BackingField = IdentifierInNewProperties,
                    AsTypeName = PropertyAsTypeName,
                    ParameterName = ParameterName,
                    GenerateLetter = ImplementLet,
                    GenerateSetter = ImplementSet,
                    UsesSetAssignment = Declaration.IsObject,
                    IsUDTProperty = true,
                    Declaration = Declaration
                };
            }
        }

    }
}
