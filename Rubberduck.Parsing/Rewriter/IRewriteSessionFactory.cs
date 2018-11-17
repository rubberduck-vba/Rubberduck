using System;

namespace Rubberduck.Parsing.Rewriter
{
    public interface IRewriteSessionFactory
    {
        IRewriteSession CodePaneSession(Func<IRewriteSession, bool> rewritingAllowed);
        IRewriteSession AttributesSession(Func<IRewriteSession, bool> rewritingAllowed);
    }
}