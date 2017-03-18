using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.PostProcessing.RewriterInfo
{
    public interface IRewriterInfoFinder
    {
        RewriterInfo GetRewriterInfo(ParserRuleContext context, Declaration target);
    }
}