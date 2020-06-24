using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Refactorings.Exceptions;

namespace Rubberduck.Refactorings
{
    public abstract class CodeOnlyRefactoringActionBase<TModel> : ICodeOnlyRefactoringAction<TModel>
        where TModel : class, IRefactoringModel
    {
        private readonly IRewritingManager _rewritingManager;

        protected CodeOnlyRefactoringActionBase(IRewritingManager rewritingManager)
        {
            _rewritingManager = rewritingManager;
        }

        public abstract void Refactor(TModel model, IRewriteSession rewriteSession);

        public virtual void Refactor(TModel model)
        {
            var rewriteSession = RewriteSession(RewriteSessionCodeKind(model));

            Refactor(model, rewriteSession);

            if (!rewriteSession.TryRewrite())
            {
                throw new RewriteFailedException(rewriteSession);
            }
        }

        private IExecutableRewriteSession RewriteSession(CodeKind codeKind)
        {
            return codeKind == CodeKind.AttributesCode
                ? _rewritingManager.CheckOutAttributesSession()
                : _rewritingManager.CheckOutCodePaneSession();
        }

        protected virtual CodeKind RewriteSessionCodeKind(TModel model) => CodeKind.CodePaneCode;
    }
}