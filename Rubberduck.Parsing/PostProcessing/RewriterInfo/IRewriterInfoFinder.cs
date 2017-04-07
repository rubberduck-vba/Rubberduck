using Antlr4.Runtime;

namespace Rubberduck.Parsing.PostProcessing.RewriterInfo
{
    public interface IRewriterInfoFinder
    {
        RewriterInfo GetRewriterInfo(ParserRuleContext context);
    }
}