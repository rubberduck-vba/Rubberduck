using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.ReplaceDeclarationIdentifier
{
    /// <summary>
    /// Supports Renaming a <c>Declaration</c> independent of its <c>IdentifierReference</c>s.  
    /// To replace <c>Declaration</c>s and its <c>IdentifierReference</c>s in a single call use <c>RenameRefactoringAction</c>
    /// </summary>
    public class ReplaceDeclarationIdentifierRefactoringAction : CodeOnlyRefactoringActionBase<ReplaceDeclarationIdentifierModel>
    {
        public ReplaceDeclarationIdentifierRefactoringAction(IRewritingManager rewritingManager)
            : base(rewritingManager) { }

        public override void Refactor(ReplaceDeclarationIdentifierModel model, IRewriteSession rewriteSession)
        {
            foreach ((Declaration target, string Name) in model.TargetNewNamePairs)
            {
                var rewriter = rewriteSession.CheckOutModuleRewriter(target.QualifiedName.QualifiedModuleName);

                if (target.Context is IIdentifierContext context)
                {
                    rewriter.Replace(context.IdentifierTokens, Name);
                }
            }
        }
    }
}
