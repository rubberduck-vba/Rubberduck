using System;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Rewriter.RewriterInfo
{
    public class ArgumentRewriterInfoFinder : RewriterInfoFinderBase
    {
        public override RewriterInfo GetRewriterInfo(ParserRuleContext context)
        {
            return GetRewriterInfo(context as VBAParser.ArgumentContext, context.Parent as VBAParser.ArgumentListContext);
        }

        private static RewriterInfo GetRewriterInfo(VBAParser.ArgumentContext arg, VBAParser.ArgumentListContext argList)
        {
            if (argList == null)
            {
                throw new ArgumentNullException(nameof(argList), @"Context is null. Expecting a VBAParser.ArgumentListContext instance.");
            }

            var items = argList.argument();
            var itemIndex = items.ToList().IndexOf(arg);

            if (items.Length == 1)
            {
                return new RewriterInfo(argList.Start.TokenIndex, argList.Stop.TokenIndex);
            }

            return GetRewriterInfoForTargetRemovedFromListStmt(arg.Start, itemIndex, argList.argument());
        }
    }
}