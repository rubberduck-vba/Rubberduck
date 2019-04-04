namespace Rubberduck.Parsing.Rewriter
{
    public interface IRewritingManager
    {
        IExecutableRewriteSession CheckOutCodePaneSession();
        IExecutableRewriteSession CheckOutAttributesSession();
        void InvalidateAllSessions();
    }
}