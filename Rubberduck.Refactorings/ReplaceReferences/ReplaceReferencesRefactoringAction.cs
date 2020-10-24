using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.ReplaceReferences
{
    /// <summary>
    /// Supports Renaming an <c>IdentifierReference</c> independent of its <c>Declaration</c>.  
    /// To replace <c>Declaration</c>s and its <c>IdentifierReference</c>s in a single call use <c>RenameRefactoringAction</c>
    /// </summary>
    public class ReplaceReferencesRefactoringAction : CodeOnlyRefactoringActionBase<ReplaceReferencesModel>
    {
        public ReplaceReferencesRefactoringAction(IRewritingManager rewritingManager)
            : base(rewritingManager)
        { }

        public override void Refactor(ReplaceReferencesModel model, IRewriteSession rewriteSession)
        {
            var replacementPairByQualifiedModuleName = model.ReferenceReplacementPairs
                .Where(pr =>
                    pr.IdentifierReference.Context.GetText() != Tokens.Me
                    && !pr.IdentifierReference.IsArrayAccess
                    && !pr.IdentifierReference.IsDefaultMemberAccess)
                .GroupBy(r => r.IdentifierReference.QualifiedModuleName);

            foreach (var replacements in replacementPairByQualifiedModuleName)
            {
                var rewriter = rewriteSession.CheckOutModuleRewriter(replacements.Key);
                foreach ((IdentifierReference identifierReference, string newIdentifier) in replacements)
                {
                    (ParserRuleContext context, string replacementName) = BuildReferenceReplacementString(identifierReference, newIdentifier, model.ModuleQualifyExternalReferences);
                    rewriter.Replace(context, replacementName);
                }
            }
        }

        private (ParserRuleContext context, string replacementName) BuildReferenceReplacementString(IdentifierReference identifierReference, string NewName, bool moduleQualify)
        {
            var replacementExpression = moduleQualify && CanBeModuleQualified(identifierReference)
                ? $"{identifierReference.Declaration.QualifiedModuleName.ComponentName}.{NewName}"
                : NewName;

            return (identifierReference.Context, replacementExpression);
        }

        private static bool CanBeModuleQualified(IdentifierReference idRef)
        {
            if (idRef.QualifiedModuleName == idRef.Declaration.QualifiedModuleName)
            {
                return false;
            }

            var isLHSOfMemberAccess =
                (idRef.Context.Parent is VBAParser.MemberAccessExprContext
                    || idRef.Context.Parent is VBAParser.WithMemberAccessExprContext)
                && !(idRef.Context == idRef.Context.Parent.GetChild(0));

            return !isLHSOfMemberAccess;
        }
    }
}
