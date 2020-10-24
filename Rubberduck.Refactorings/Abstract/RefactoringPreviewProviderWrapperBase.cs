using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings
{
    public abstract class RefactoringPreviewProviderWrapperBase<TModel> : IRefactoringPreviewProvider<TModel>
        where TModel : class, IRefactoringModel
    {
        private readonly IRewritingManager _rewritingManager;
        private readonly ICodeOnlyRefactoringAction<TModel> _refactoringAction;

        protected RefactoringPreviewProviderWrapperBase(
            ICodeOnlyRefactoringAction<TModel> refactoringAction,
            IRewritingManager rewritingManager)
        {
            _refactoringAction = refactoringAction;
            _rewritingManager = rewritingManager;
        }

        protected abstract QualifiedModuleName ComponentToShow(TModel model);

        public virtual string Preview(TModel model)
        {
            var rewriteSession = RewriteSession(RewriteSessionCodeKind);
            _refactoringAction.Refactor(model, rewriteSession);
            var componentToShow = ComponentToShow(model);
            var rewriter = rewriteSession.CheckOutModuleRewriter(componentToShow);
            return rewriter.GetText();
        }

        private IExecutableRewriteSession RewriteSession(CodeKind codeKind)
        {
            return codeKind == CodeKind.AttributesCode
                ? _rewritingManager.CheckOutAttributesSession()
                : _rewritingManager.CheckOutCodePaneSession();
        }

        protected virtual CodeKind RewriteSessionCodeKind => CodeKind.CodePaneCode;
    }
}