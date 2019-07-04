namespace Rubberduck.Parsing.Rewriter
{
    public interface IExecutableRewriteSession : IRewriteSession
    {
        bool TryRewrite();
    }
}