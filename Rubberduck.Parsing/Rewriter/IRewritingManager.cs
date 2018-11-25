namespace Rubberduck.Parsing.Rewriter
{
    public interface IRewritingManager
    {
        IRewriteSession CheckOutCodePaneSession();
        IRewriteSession CheckOutAttributesSession();
        void InvalidateAllSessions();
    }
}