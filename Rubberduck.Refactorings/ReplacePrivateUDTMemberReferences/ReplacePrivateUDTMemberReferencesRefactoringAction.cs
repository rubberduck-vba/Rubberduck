using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.ReplacePrivateUDTMemberReferences
{
    /// <summary>
    /// Replaces UserDefinedTypeMember <c>IdentifierReference</c>s of a Private <c>UserDefinedType</c>
    /// with an accessor expression.
    /// </summary>
    public class ReplacePrivateUDTMemberReferencesRefactoringAction : CodeOnlyRefactoringActionBase<ReplacePrivateUDTMemberReferencesModel>
    {
        private Dictionary<IdentifierReference, (ParserRuleContext, string)> IdentifierReplacements { get; } = new Dictionary<IdentifierReference, (ParserRuleContext, string)>();

        public ReplacePrivateUDTMemberReferencesRefactoringAction(IRewritingManager rewritingManager)
            : base(rewritingManager) { }

        public override void Refactor(ReplacePrivateUDTMemberReferencesModel model, IRewriteSession rewriteSession)
        {
            if (!(model.UDTMembers?.Any() ?? false))
            {
                return;
            }

            SetRewriteContent(model);

            RewriteReferences(rewriteSession);
        }

        private void SetRewriteContent(ReplacePrivateUDTMemberReferencesModel model)
        {
            var targetReferences = model.Targets.Select(t => model.UserDefinedTypeInstance(t))
                .SelectMany(udtInstance => udtInstance.UDTMemberReferences);

            foreach (var idRef in targetReferences)
            {
                if (model.TryGetLocalReferenceExpression(idRef, out var expression))
                {
                    SetUDTMemberReferenceReplacementExpression(idRef, expression);
                }
            }
        }

        //TODO: Write tests using Member and With access contexts locally and externally - replacement text is different for readOnly scenario
        private void SetUDTMemberReferenceReplacementExpression(IdentifierReference idRef, string replacementText, bool moduleQualify = false)
        {
            idRef.Context.TryGetAncestor<VBAParser.MemberAccessExprContext>(out var maec); 
            idRef.Context.TryGetAncestor<VBAParser.WithMemberAccessExprContext>(out var wmac);

            if (maec is null && wmac is null)
            {
                throw new ArgumentException();
            }

            if (maec != null)
            {
                if (maec.TryGetChildContext<VBAParser.WithMemberAccessExprContext>(out _))
                {
                    replacementText = $".{replacementText}";
                }
                AddIdentifierReplacement(idRef, maec, replacementText);
            }
            else
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
