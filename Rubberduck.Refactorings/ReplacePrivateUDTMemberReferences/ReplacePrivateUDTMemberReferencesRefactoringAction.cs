using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.ReplacePrivateUDTMemberReferences
{
    /// <summary>
    /// Replaces UserDefinedTypeMember <c>IdentifierReference</c>s of a Private <c>UserDefinedType</c>
    /// with a Property <c>IdentifierReference</c>.
    /// </summary>
    public class ReplacePrivateUDTMemberReferencesRefactoringAction : CodeOnlyRefactoringActionBase<ReplacePrivateUDTMemberReferencesModel>
    {
        private Dictionary<IdentifierReference, (ParserRuleContext, string)> IdentifierReplacements { get; } = new Dictionary<IdentifierReference, (ParserRuleContext, string)>();

        public ReplacePrivateUDTMemberReferencesRefactoringAction(IRewritingManager rewritingManager)
            : base(rewritingManager)
        { }

        public override void Refactor(ReplacePrivateUDTMemberReferencesModel model, IRewriteSession rewriteSession)
        {
            if (!(model.UDTMembers?.Any() ?? false))
            {
                return;
            }

            foreach (var target in model.Targets)
            {
                SetRewriteContent(target, model);
            }

            RewriteReferences(rewriteSession);
        }

        private void SetRewriteContent(VariableDeclaration target, ReplacePrivateUDTMemberReferencesModel model)
        {
            var udtInstance = model.UserDefinedTypeInstance(target);
            foreach (var idRef in udtInstance.UDTMemberReferences)
            {
                var internalExpression = model.LocalReferenceExpression(target, idRef.Declaration);
                if (internalExpression.HasValue)
                {
                    SetUDTMemberReferenceRewriteContent(target, idRef, internalExpression.Expression);
                }
            }
        }

        private void SetUDTMemberReferenceRewriteContent(VariableDeclaration instanceField, IdentifierReference idRef, string replacementText, bool moduleQualify = false)
        {
            if (idRef.Context.TryGetAncestor<VBAParser.MemberAccessExprContext>(out var maec))
            {
                if (maec.TryGetChildContext<VBAParser.MemberAccessExprContext>(out var childMaec))
                {
                    if (childMaec.TryGetChildContext<VBAParser.SimpleNameExprContext>(out var smp))
                    {
                        AddIdentifierReplacement(idRef, maec, $"{smp.GetText()}.{replacementText}");
                    }
                }
                else if (maec.TryGetChildContext<VBAParser.WithMemberAccessExprContext>(out var wm))
                {
                    AddIdentifierReplacement(idRef, maec, $".{replacementText}");
                }
                else
                {
                    AddIdentifierReplacement(idRef, maec, replacementText);
                }
            }
            else if (idRef.Context.TryGetAncestor<VBAParser.WithMemberAccessExprContext>(out var wmac))
            {
                AddIdentifierReplacement(idRef, wmac, replacementText);
            }
        }

        private void AddIdentifierReplacement(IdentifierReference idRef, ParserRuleContext context, string replacementText)
        {
            if (IdentifierReplacements.ContainsKey(idRef))
            {
                IdentifierReplacements[idRef] = (context, replacementText);
                return;
            }
            IdentifierReplacements.Add(idRef, (context, replacementText));
        }

        private void RewriteReferences(IRewriteSession rewriteSession)
        {
            foreach (var replacement in IdentifierReplacements)
            {
                (ParserRuleContext Context, string Text) = replacement.Value;
                var rewriter = rewriteSession.CheckOutModuleRewriter(replacement.Key.QualifiedModuleName);
                rewriter.Replace(Context, Text);
            }
        }
    }
}
