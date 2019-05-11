using System;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing.Rewriter
{
    public class RewriteSessionFactory : IRewriteSessionFactory
    {
        private readonly RubberduckParserState _state;
        private readonly IRewriterProvider _rewriterProvider;
        private readonly ISelectionRecoverer _selectionRecoverer;

        public RewriteSessionFactory(RubberduckParserState state, IRewriterProvider rewriterProvider, ISelectionRecoverer selectionRecoverer)
        {
            _state = state;
            _rewriterProvider = rewriterProvider;
            _selectionRecoverer = selectionRecoverer;
        }

        public IExecutableRewriteSession CodePaneSession(Func<IRewriteSession, bool> rewritingAllowed)
        {
            return new CodePaneRewriteSession(_state, _rewriterProvider, _selectionRecoverer, rewritingAllowed);
        }

        public IExecutableRewriteSession AttributesSession(Func<IRewriteSession, bool> rewritingAllowed)
        {
            return new AttributesRewriteSession(_state, _rewriterProvider, _selectionRecoverer, rewritingAllowed);
        }
    }
}