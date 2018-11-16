using System;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing.Rewriter
{
    public class RewriteSessionFactory : IRewriteSessionFactory
    {
        private readonly RubberduckParserState _state;
        private readonly IRewriterProvider _rewriterProvider;

        public RewriteSessionFactory(RubberduckParserState state, IRewriterProvider rewriterProvider)
        {
            _state = state;
            _rewriterProvider = rewriterProvider;
        }

        public IRewriteSession CodePaneSession(Func<IRewriteSession, bool> rewritingAllowed)
        {
            return new CodePaneRewriteSession(_state, _rewriterProvider, rewritingAllowed);
        }

        public IRewriteSession AttributesSession(Func<IRewriteSession, bool> rewritingAllowed)
        {
            return new AttributesRewriteSession(_state, _rewriterProvider, rewritingAllowed);
        }
    }
}