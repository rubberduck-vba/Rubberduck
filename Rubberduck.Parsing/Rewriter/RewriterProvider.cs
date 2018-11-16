using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Rewriter
{
    public class RewriterProvider : IRewriterProvider
    {
        private ITokenStreamCache _tokenStreamCache;
        private IModuleRewriterFactory _rewriterFactory;

        public RewriterProvider(ITokenStreamCache tokenStreamCache, IModuleRewriterFactory rewriterFactory)
        {
            _tokenStreamCache = tokenStreamCache;
            _rewriterFactory = rewriterFactory;
        }


        public IExecutableModuleRewriter CodePaneModuleRewriter(QualifiedModuleName module)
        {
            var tokenStream = _tokenStreamCache.CodePaneTokenStream(module);
            return _rewriterFactory.CodePaneRewriter(module, tokenStream);
        }

        public IExecutableModuleRewriter AttributesModuleRewriter(QualifiedModuleName module)
        {
            var tokenStream = _tokenStreamCache.AttributesTokenStream(module);
            return _rewriterFactory.AttributesRewriter(module, tokenStream);
        }
    }
}