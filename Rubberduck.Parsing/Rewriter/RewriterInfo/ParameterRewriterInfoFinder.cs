using System;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Rewriter.RewriterInfo
{
    public class ParameterRewriterInfoFinder : RewriterInfoFinderBase
    {
        public override RewriterInfo GetRewriterInfo(ParserRuleContext context)
        {
            return GetRewriterInfo(context as VBAParser.ArgContext, context.Parent as VBAParser.ArgListContext);
        }

        private static RewriterInfo GetRewriterInfo(VBAParser.ArgContext arg, VBAParser.ArgListContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context), @"Context is null. Expecting a VBAParser.ArgListContext instance.");
            }

            var items = context.arg();
            var itemIndex = items.ToList().IndexOf(arg);

            if (items.Length == 1)
            {
                return new RewriterInfo(context.LPAREN().Symbol.TokenIndex + 1, context.RPAREN().Symbol.TokenIndex - 1);
            }

            return GetRewriterInfoForTargetRemovedFromListStmt(arg.Start, itemIndex, context.arg());
        }
    }
}