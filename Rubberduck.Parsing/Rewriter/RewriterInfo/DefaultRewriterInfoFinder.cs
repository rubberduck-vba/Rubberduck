using Antlr4.Runtime;

namespace Rubberduck.Parsing.Rewriter.RewriterInfo
{
    public class DefaultRewriterInfoFinder : RewriterInfoFinderBase
    {
        public override RewriterInfo GetRewriterInfo(ParserRuleContext context)
        {
            return new RewriterInfo(context.Start.TokenIndex, context.Stop.TokenIndex);
        }
    }
}