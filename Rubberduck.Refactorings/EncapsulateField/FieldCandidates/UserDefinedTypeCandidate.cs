using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IUserDefinedTypeCandidate : IEncapsulateFieldCandidate
    {
        IEnumerable<IUserDefinedTypeMemberCandidate> Members { get; }
        void AddMember(IUserDefinedTypeMemberCandidate member);
        bool TypeDeclarationIsPrivate { set; get; }
    }

    public class UserDefinedTypeCandidate : EncapsulateFieldCandidate, IUserDefinedTypeCandidate
    {
        public UserDefinedTypeCandidate(Declaration declaration, IEncapsulateFieldNamesValidator validator)
            : base(declaration, validator)
        {
            PropertyAccessor = AccessorTokens.Field;
            ReferenceAccessor = AccessorTokens.Field;
        }

        public void AddMember(IUserDefinedTypeMemberCandidate member)
        {
            _udtMembers.Add(member);
        }

        private List<IUserDefinedTypeMemberCandidate> _udtMembers = new List<IUserDefinedTypeMemberCandidate>();
        public IEnumerable<IUserDefinedTypeMemberCandidate> Members => _udtMembers;

        public bool TypeDeclarationIsPrivate { set; get; }

        public override string FieldIdentifier
        {
            get => TypeDeclarationIsPrivate ? _fieldAndProperty.TargetFieldName : _fieldAndProperty.Field;
            set => _fieldAndProperty.Field = value;
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

                        //Reaching this line typically implies that there are multiple fields of the same Type within the module.
                        //Try to use a name involving the parent's identifier to make it unique/meaningful 
                        //before appending incremented value(s).
                        member.PropertyName = $"{FieldIdentifier.Capitalize()}{member.PropertyName.Capitalize()}";
                        _validator.AssignNoConflictIdentifier(member, DeclarationType.Property);
                    }
                }
                base.EncapsulateFlag = value;
            }
            get => _encapsulateFlag;
        }


        public override string ReferenceQualifier
        {
            set
            {
                _referenceQualifier = value;
                PropertyAccessor = (value?.Length ?? 0) == 0
                    ? AccessorTokens.Field
                    : AccessorTokens.Property;
            }
            get => _referenceQualifier;
        }

        //public bool FieldQualifyUDTMemberPropertyNames
        //{
        //    set
        //    {
        //        foreach (var member in Members)
        //        {
        //            member.IncludeParentNameWithPropertyIdentifier = value;
        //        }
        //    }

        //    get => Members.All(m => m.IncludeParentNameWithPropertyIdentifier);
        //}

        protected override void LoadFieldReferenceContextReplacements()
        {
            if (TypeDeclarationIsPrivate)
            {
                LoadPrivateUDTFieldReferenceExpressions();
                LoadUDTMemberReferenceExpressions();
                return;
            }
            base.LoadFieldReferenceContextReplacements();
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

        public override void StageFieldReferenceReplacements(IStateUDT stateUDT = null)
        {

            PropertyAccessor = stateUDT is null ? AccessorTokens.Field : AccessorTokens.Property;
            ReferenceAccessor = AccessorTokens.Property;
            ReferenceQualifier = stateUDT?.FieldIdentifier ?? null;
            LoadFieldReferenceContextReplacements();
            foreach (var member in Members)
            {
                member.StageFieldReferenceReplacements(stateUDT);
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
                //if (!_validator.IsSelfConsistent(this, out errorMessage))
                //{
                    return false;
                //}
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

        private void LoadPrivateUDTFieldReferenceExpressions()
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
                foreach (var rf in member.FieldRelatedReferences(this))
                {
                    if (rf.QualifiedModuleName == QualifiedModuleName
                        && !rf.Context.TryGetAncestor<VBAParser.WithMemberAccessExprContext>(out _))
                    {
                        member.SetReferenceRewriteContent(rf, member.PropertyName);
                    }
                    else
                    {
                        var moduleQualifier = rf.Context.TryGetAncestor<VBAParser.WithStmtContext>(out _)
                            || rf.QualifiedModuleName == QualifiedModuleName
                            ? string.Empty
                            : $"{QualifiedModuleName.ComponentName}";

                        member.SetReferenceRewriteContent(rf, $"{moduleQualifier}.{member.PropertyName}");
                    }
                }
            }
        }
    }
}
