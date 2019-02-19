namespace Rubberduck.Parsing.Rewriter
{
    public interface IMemberAttributeRecovererWithSettableRewritingManager : IMemberAttributeRecoverer
    {
        IRewritingManager RewritingManager { set; }
    }
}