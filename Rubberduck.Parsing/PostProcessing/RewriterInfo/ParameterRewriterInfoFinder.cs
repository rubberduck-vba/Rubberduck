using System;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.PostProcessing.RewriterInfo
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
            var count = items.Count;

            if (count == 1)
            {
                return new RewriterInfo(context.LPAREN().Symbol.TokenIndex + 1, context.RPAREN().Symbol.TokenIndex - 1);
            }

            return GetRewriterInfoForTargetRemovedFromListStmt(arg.Start, itemIndex, context.arg());
        }
    }
}