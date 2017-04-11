using Antlr4.Runtime;

namespace Rubberduck.Parsing.Rewriter.RewriterInfo
{
    public interface IRewriterInfoFinder
    {
        RewriterInfo GetRewriterInfo(ParserRuleContext context);
    }
}