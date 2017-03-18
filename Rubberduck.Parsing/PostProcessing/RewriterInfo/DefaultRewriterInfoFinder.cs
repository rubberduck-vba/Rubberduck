using System.Diagnostics;
using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.PostProcessing.RewriterInfo
{
    public class DefaultRewriterInfoFinder : RewriterInfoFinderBase
    {
        public override RewriterInfo GetRewriterInfo(ParserRuleContext context, Declaration target)
        {
            Debug.Assert(target.Context == context);
            return new RewriterInfo(context.Start.TokenIndex, context.Stop.TokenIndex);
        }
    }
}