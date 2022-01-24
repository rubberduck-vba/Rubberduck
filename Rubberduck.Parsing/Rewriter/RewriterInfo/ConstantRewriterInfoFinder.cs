using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Rewriter.RewriterInfo
{
    public class ConstantRewriterInfoFinder : RewriterInfoFinderBase
    {
        public override RewriterInfo GetRewriterInfo(ParserRuleContext context)
        {
            return GetRewriterInfo(context as VBAParser.ConstSubStmtContext, context.Parent as VBAParser.ConstStmtContext);
        }

        private static RewriterInfo GetRewriterInfo(VBAParser.ConstSubStmtContext target, VBAParser.ConstStmtContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context), @"Context is null. Expecting a VBAParser.ConstStmtContext instance.");
            }

            var items = context.constSubStmt();
            var itemIndex = items.ToList().IndexOf(target);
            var count = items.Length;

            var element = context.Parent?.Parent as VBAParser.ModuleDeclarationsElementContext;
            if (element != null)
            {
                return GetModuleConstantRemovalInfo(target, element, count, itemIndex, items);
            }

            return GetLocalConstantRemovalInfo(target, context, count, itemIndex, items);
        }

        private static RewriterInfo GetModuleConstantRemovalInfo(
            VBAParser.ConstSubStmtContext target, 
            VBAParser.ModuleDeclarationsElementContext element,
            int count, 
            int itemIndex, 
            IReadOnlyList<VBAParser.ConstSubStmtContext> items)
        {
            if (count == 1)
            {
                return GetSeparateModuleConstantRemovalInfo(element);
            }
            return GetRewriterInfoForTargetRemovedFromListStmt(target.Start, itemIndex, items);
        }

        private static RewriterInfo GetSeparateModuleConstantRemovalInfo(VBAParser.ModuleDeclarationsElementContext element)
        {
            var startIndex = element.Start.TokenIndex;
            var stopIndex = FindStopTokenIndexForRemoval(element);
            return new RewriterInfo(startIndex, stopIndex);
        }

        private static RewriterInfo GetLocalConstantRemovalInfo(VBAParser.ConstSubStmtContext target,
            VBAParser.ConstStmtContext constants,
            int count, int itemIndex, IReadOnlyList<VBAParser.ConstSubStmtContext> items)
        {
            if (count == 1)
            {
                return GetSeparateLocalConstantRemovalInfo(constants);
            }
            return GetRewriterInfoForTargetRemovedFromListStmt(target.Start, itemIndex, items);
        }

        private static RewriterInfo GetSeparateLocalConstantRemovalInfo(VBAParser.ConstStmtContext constStmtContext)
        {
            var mainBlockStmt = constStmtContext.GetAncestor<VBAParser.MainBlockStmtContext>();
            var startIndex = mainBlockStmt.Start.TokenIndex;

            var stopIndex = FindStopTokenIndexForRemoval(mainBlockStmt);

            return new RewriterInfo(startIndex, stopIndex);
        }
    }
}