using System;

namespace Rubberduck.Parsing.Rewriter
{
    public interface IRewriteSessionFactory
    {
        IExecutableRewriteSession CodePaneSession(Func<IRewriteSession, bool> rewritingAllowed);
        IExecutableRewriteSession AttributesSession(Func<IRewriteSession, bool> rewritingAllowed);
    }
}