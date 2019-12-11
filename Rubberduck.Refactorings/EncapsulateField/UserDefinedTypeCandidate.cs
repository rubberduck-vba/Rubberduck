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
        bool FieldQualifyUDTMemberPropertyName { set; get; }
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

        public override string NewFieldName
        {
            get => TypeDeclarationIsPrivate ? _fieldAndProperty.TargetFieldName : _fieldAndProperty.Field;
            set => _fieldAndProperty.Field = value;
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

        public bool FieldQualifyUDTMemberPropertyName
        {
            set
            {
                foreach (var member in Members)
                {
                    member.FieldQualifyUDTMemberPropertyName = value;
                }
            }

            get => Members.All(m => m.FieldQualifyUDTMemberPropertyName);
        }

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
                return new List<IPropertyGeneratorAttributes>() { AsPropertyGeneratorSpec };
            }
        }

        public override void SetupReferenceReplacements(IStateUDTField stateUDT = null)
        {

            PropertyAccessor = stateUDT is null ? AccessorTokens.Field : AccessorTokens.Property;
            ReferenceAccessor = AccessorTokens.Property;
            ReferenceQualifier = stateUDT?.NewFieldName ?? null;
            LoadFieldReferenceContextReplacements();
            foreach (var member in Members)
            {
                member.SetupReferenceReplacements(stateUDT);
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

        private void LoadPrivateUDTFieldReferenceExpressions()
        {
            foreach (var idRef in Declaration.References)
            {
                if (idRef.QualifiedModuleName == QualifiedModuleName
                    && idRef.Context.Parent.Parent is VBAParser.WithStmtContext wsc)
                {
                    SetReferenceRewriteContent(idRef, NewFieldName);
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
                        //If rf is a WithMemberAccess expression, modify the LExpr.  e.g. "With this" => "With <qmn.ModuleName>"
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
