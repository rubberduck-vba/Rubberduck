using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IUserDefinedTypeCandidate : IEncapsulateFieldCandidate
    {
        IEnumerable<IUserDefinedTypeMemberCandidate> Members { get; }
        void AddMember(IUserDefinedTypeMemberCandidate member);
        bool TypeDeclarationIsPrivate { set; get; }
        bool CanBeObjectStateUDT { set; get; }
    }

    public class UserDefinedTypeCandidate : EncapsulateFieldCandidate, IUserDefinedTypeCandidate
    {
        public UserDefinedTypeCandidate(Declaration declaration, IValidateEncapsulateFieldNames validator)
            : base(declaration, validator)
        {
            NewPropertyAccessor = AccessorMember.Field;
            ReferenceAccessor = AccessorMember.Field;
        }

        public void AddMember(IUserDefinedTypeMemberCandidate member)
        {
            _udtMembers.Add(member);
        }

        private List<IUserDefinedTypeMemberCandidate> _udtMembers = new List<IUserDefinedTypeMemberCandidate>();
        public IEnumerable<IUserDefinedTypeMemberCandidate> Members => _udtMembers;

        public bool TypeDeclarationIsPrivate { set; get; }

        private bool _canBeObjectStateUDT;
        public bool CanBeObjectStateUDT
        {
            set => _canBeObjectStateUDT = value;
            get => _canBeObjectStateUDT;
        }

        public override string FieldIdentifier
        {
            get => TypeDeclarationIsPrivate ? _fieldAndProperty.TargetFieldName : _fieldAndProperty.Field;
            set => _fieldAndProperty.Field = value;
        }

        private string _referenceQualifier;
        public override string ReferenceQualifier
        {
            set
            {
                _referenceQualifier = value;
                foreach( var member in Members)
                {
                    member.ReferenceQualifier = ReferenceWithinNewProperty;
                }
            }
            get => _referenceQualifier;
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
                        if (!_validator.HasConflictingIdentifier(member, DeclarationType.Property, out _))
                        {
                            continue;
                        }

                        //Reaching this line probably implies that there are multiple fields of the same User Defined 
                        //Type within the module.
                        //Try to use a name involving the parent's identifier to make it unique/meaningful 
                        //before giving up and creating incremented value(s).
                        member.PropertyName = $"{FieldIdentifier.CapitalizeFirstLetter()}{member.PropertyName.CapitalizeFirstLetter()}";
                        _validator.AssignNoConflictIdentifier(member, DeclarationType.Property);
                    }
                }
                base.EncapsulateFlag = value;
            }
            get => _encapsulateFlag;
        }

        public override void LoadFieldReferenceContextReplacements()
        {
            if (TypeDeclarationIsPrivate)
            {
                LoadPrivateUDTFieldLocalReferenceExpressions();
                LoadUDTMemberReferenceExpressions();
                return;
            }

            foreach (var idRef in Declaration.References)
            {
                var replacementText = RequiresAccessQualification(idRef)
                    ? $"{QualifiedModuleName.ComponentName}.{PropertyName}"
                    : PropertyName;

                SetReferenceRewriteContent(idRef, replacementText);
            }
        }

        public override IEnumerable<IPropertyGeneratorAttributes> PropertyAttributeSets
        {
            get
            {
                if (TypeDeclarationIsPrivate)
                {
                    var specs = new List<IPropertyGeneratorAttributes>();
                    foreach (var member in Members)
                    {
                        specs.Add(member.AsPropertyGeneratorSpec);
                    }
                    return specs;
                }
                return new List<IPropertyGeneratorAttributes>() { AsPropertyAttributeSet };
            }
        }

        public override IEnumerable<KeyValuePair<IdentifierReference, (ParserRuleContext, string)>> ReferenceReplacements
        {
            get
            {
                var results = new List<KeyValuePair<IdentifierReference, (ParserRuleContext, string)>>();
                foreach (var replacement in IdentifierReplacements)
                {
                    var kv = new KeyValuePair<IdentifierReference, (ParserRuleContext, string)>
                        (replacement.Key, replacement.Value);
                    results.Add(kv);
                }

                foreach (var replacement in Members.SelectMany(m => m.IdentifierReplacements))
                {
                    var kv = new KeyValuePair<IdentifierReference, (ParserRuleContext, string)>
                        (replacement.Key, replacement.Value);
                    results.Add(kv);
                }
                return results;
            }
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

        public override bool TryValidateEncapsulationAttributes(out string errorMessage)
        {
            errorMessage = string.Empty;
            if (!EncapsulateFlag) { return true; }

            if (!_validator.IsValidVBAIdentifier(PropertyName, DeclarationType.Property, out errorMessage))
            {
                return false;
            }

            if (!TypeDeclarationIsPrivate && !_validator.IsSelfConsistent(this, out errorMessage))
            {
                return false;
            }

            if (_validator.HasConflictingIdentifier(this, DeclarationType.Property, out errorMessage))
            {
                return false;
            }

            if (_validator.HasConflictingIdentifier(this, DeclarationType.Variable, out errorMessage))
            {
                return false;
            }
            return true;
        }

        protected override IPropertyGeneratorAttributes AsPropertyAttributeSet
        {
            get
            {
                return new PropertyAttributeSet()
                {
                    PropertyName = PropertyName,
                    BackingField = ReferenceWithinNewProperty,
                    AsTypeName = AsTypeName_Property,
                    ParameterName = ParameterName,
                    GenerateLetter = ImplementLet,
                    GenerateSetter = ImplementSet,
                    UsesSetAssignment = Declaration.IsObject,
                    IsUDTProperty = true
                };
            }
        }

        private void LoadPrivateUDTFieldLocalReferenceExpressions()
        {
            foreach (var idRef in Declaration.References)
            {
                if (idRef.QualifiedModuleName == QualifiedModuleName
                    && idRef.Context.Parent.Parent is VBAParser.WithStmtContext wsc)
                {
                    SetReferenceRewriteContent(idRef, FieldIdentifier);
                }
            }
        }

        private void LoadUDTMemberReferenceExpressions()
        {
            foreach (var member in Members)
            {
                member.LoadReferenceExpressions();
            }
        }
    }
}
