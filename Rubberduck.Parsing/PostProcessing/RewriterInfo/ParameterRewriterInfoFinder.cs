using System;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.PostProcessing.RewriterInfo
{
    public class ParameterRewriterInfoFinder : RewriterInfoFinderBase
    {
        public override RewriterInfo GetRewriterInfo(ParserRuleContext context, Declaration target)
        {
            return GetRewriterInfo(target, context.Parent as VBAParser.ArgListContext);
        }

        private static RewriterInfo GetRewriterInfo(Declaration target, VBAParser.ArgListContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException("context", @"Context is null. Expecting a VBAParser.ArgListContext instance.");
            }

            var items = context.arg();
            var itemIndex = items.ToList().IndexOf((VBAParser.ArgContext)target.Context);
            var count = items.Count;

            if (count == 1)
            {
                return new RewriterInfo(context.LPAREN().Symbol.TokenIndex + 1, context.RPAREN().Symbol.TokenIndex - 1);
            }
            return GetRewriterInfoForTargetRemovedFromListStmt(target.Context.Start, itemIndex, context.arg());
        }
    }
}