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
    public interface IEncapsulatedUserDefinedTypeMember : IEncapsulateFieldCandidate
    {
        IEncapsulatedUserDefinedTypeField Parent { get; }
        bool FieldQualifyPropertyName { set; get; }
        IPropertyGeneratorSpecification AsPropertyGeneratorSpec { get; }
        Dictionary<IdentifierReference, RewriteReplacePair> IdentifierReplacements { get; }
    }

    public class EncapsulatedUserDefinedTypeMember : EncapsulateFieldCandidate, IEncapsulatedUserDefinedTypeMember
    {
        public EncapsulatedUserDefinedTypeMember(Declaration target, IEncapsulatedUserDefinedTypeField udtVariable, IEncapsulateFieldNamesValidator validator)
            : base(target, validator)
        {
            Parent = udtVariable;

            PropertyName = IdentifierName;
            PropertyAccessExpression = () => $"{Parent.PropertyAccessExpression()}.{PropertyName}";
            ReferenceExpression = () => $"{Parent.PropertyAccessExpression()}.{PropertyName}";
        }

        public IEncapsulatedUserDefinedTypeField Parent { private set; get; }

        private bool _fieldNameQualifyProperty;
        public bool FieldQualifyPropertyName
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

        public override string TargetID => $"{Parent.IdentifierName}.{IdentifierName}";

        public override IEnumerable<IdentifierReference> References
        {
            get
            {
                //var references = new List<IdentifierReference>();
                //foreach (var member in Members)
                //{
                //references.AddRange(GetUDTMemberReferencesForField(this, Parent));
                //}
                return GetUDTMemberReferencesForField(this, Parent);
            }
        }

        public override void AddReferenceReplacement(IdentifierReference idRef, string replacementText)
        {
            Debug.Assert(idRef.Context.Parent is ParserRuleContext, "idRef.Context.Parent is not convertable to ParserRuleContext");
            //if (idRef.Context.Parent is ParserRuleContext prContext)
            //{
                IdentifierReplacements.Add(idRef, new RewriteReplacePair(replacementText, idRef.Context.Parent as ParserRuleContext));
                return;
            //}
        }

        public new IPropertyGeneratorSpecification AsPropertyGeneratorSpec
            => base.AsPropertyGeneratorSpec;

        public new Dictionary<IdentifierReference, RewriteReplacePair> IdentifierReplacements { get; } = new Dictionary<IdentifierReference, RewriteReplacePair>();


        private IEnumerable<IdentifierReference> GetUDTMemberReferencesForField(IEncapsulateFieldCandidate udtMember, IEncapsulatedUserDefinedTypeField field)
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

        //private bool RequiresAccessQualification(IdentifierReference idRef)
        //{
        //    var isLHSOfMemberAccess =
        //                (idRef.Context.Parent is VBAParser.MemberAccessExprContext
        //                    || idRef.Context.Parent is VBAParser.WithMemberAccessExprContext)
        //                && !(idRef.Context == idRef.Context.Parent.GetChild(0));// is VBAParser.SimpleNameExprContext))

        //    return idRef.QualifiedModuleName != idRef.Declaration.QualifiedModuleName
        //                && !isLHSOfMemberAccess;
        //}
    }
}
