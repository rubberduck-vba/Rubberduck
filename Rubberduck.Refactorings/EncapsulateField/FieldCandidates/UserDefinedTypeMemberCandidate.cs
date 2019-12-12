using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IUserDefinedTypeMemberCandidate : IEncapsulateFieldCandidate
    {
        IUserDefinedTypeCandidate Parent { get; }
        bool FieldQualifyUDTMemberPropertyName { set; get; }
        IPropertyGeneratorAttributes AsPropertyGeneratorSpec { get; }
        Dictionary<IdentifierReference, (ParserRuleContext, string)> IdentifierReplacements { get; }
        IEnumerable<IdentifierReference> FieldRelatedReferences(IUserDefinedTypeCandidate field);
    }

    public class UserDefinedTypeMemberCandidate : EncapsulateFieldCandidate, IUserDefinedTypeMemberCandidate
    {
        public UserDefinedTypeMemberCandidate(Declaration target, IUserDefinedTypeCandidate udtVariable, IEncapsulateFieldNamesValidator validator)
            : base(target, validator)
        {
            Parent = udtVariable;

            PropertyName = IdentifierName;
            PropertyAccessor = AccessorTokens.Property;
            ReferenceAccessor = AccessorTokens.Property;
        }

        public IUserDefinedTypeCandidate Parent { private set; get; }

        private bool _fieldNameQualifyProperty;
        public bool FieldQualifyUDTMemberPropertyName
        {
            get => _fieldNameQualifyProperty;
            set
            {
                _fieldNameQualifyProperty = value;
                PropertyName = _fieldNameQualifyProperty
                    ? $"{Parent.IdentifierName.Capitalize()}_{IdentifierName}"
                    : IdentifierName;
            }
        }

        public override void StageFieldReferenceReplacements(IStateUDT stateUDT = null) { }

        public override string ReferenceQualifier
        {
            set => _referenceQualifier = value;
            get => Parent.ReferenceWithinNewProperty;
        }

        public override string TargetID => $"{Parent.IdentifierName}.{IdentifierName}";

        public IEnumerable<IdentifierReference> FieldRelatedReferences(IUserDefinedTypeCandidate field)
            => GetUDTMemberReferencesForField(this, field);

        public override void SetReferenceRewriteContent(IdentifierReference idRef, string replacementText)
        {
            Debug.Assert(idRef.Context.Parent is ParserRuleContext, "idRef.Context.Parent is not convertable to ParserRuleContext");

            if (IdentifierReplacements.ContainsKey(idRef))
            {
                IdentifierReplacements[idRef] = (idRef.Context.Parent as ParserRuleContext, replacementText);
                return;
            }
            IdentifierReplacements.Add(idRef, (idRef.Context.Parent as ParserRuleContext, replacementText));
        }

        public new IPropertyGeneratorAttributes AsPropertyGeneratorSpec
        {
            get
            {
                if (_fieldNameQualifyProperty)
                {
                    PropertyAccessor = AccessorTokens.Field;
                }

                return new PropertyAttributeSet()
                {
                    PropertyName = PropertyName,
                    BackingField = ReferenceWithinNewProperty,
                    AsTypeName = AsTypeName,
                    ParameterName = ParameterName,
                    GenerateLetter = ImplementLetSetterType,
                    GenerateSetter = ImplementSetSetterType,
                    UsesSetAssignment = Declaration.IsObject
                };
            }
        }

        public new Dictionary<IdentifierReference, (ParserRuleContext, string)> IdentifierReplacements { get; } = new Dictionary<IdentifierReference, (ParserRuleContext, string)>();

        private IEnumerable<IdentifierReference> GetUDTMemberReferencesForField(IEncapsulateFieldCandidate udtMember, IUserDefinedTypeCandidate field)
        {
            var refs = new List<IdentifierReference>();
            foreach (var idRef in udtMember.Declaration.References)
            {
                if (idRef.Context.TryGetAncestor<VBAParser.MemberAccessExprContext>(out var mac))
                {
                    var LHS = mac.children.First();
                    switch (LHS)
                    {
                        case VBAParser.SimpleNameExprContext snec:
                            if (snec.GetText().Equals(field.IdentifierName))
                            {
                                refs.Add(idRef);
                            }
                            break;
                        case VBAParser.MemberAccessExprContext submac:
                            if (submac.children.Last() is VBAParser.UnrestrictedIdentifierContext ur && ur.GetText().Equals(field.IdentifierName))
                            {
                                refs.Add(idRef);
                            }
                            break;
                        case VBAParser.WithMemberAccessExprContext wmac:
                            if (wmac.children.Last().GetText().Equals(field.IdentifierName))
                            {
                                refs.Add(idRef);
                            }
                            break;
                    }
                }
                else if (idRef.Context.TryGetAncestor<VBAParser.WithMemberAccessExprContext>(out var wmac))
                {
                    var wm = wmac.GetAncestor<VBAParser.WithStmtContext>();
                    var Lexpr = wm.GetChild<VBAParser.LExprContext>();
                    if (Lexpr.GetText().Equals(field.IdentifierName))
                    {
                        refs.Add(idRef);
                    }
                }
            }
            return refs;
        }
    }
}
