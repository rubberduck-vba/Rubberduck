using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.PostProcessing.RewriterInfo
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
            var count = items.Count;

            var element = context.Parent as VBAParser.ModuleDeclarationsElementContext;
            if (element != null)
            {
                return GetModuleConstantRemovalInfo(target, element, count, itemIndex, items);
            }

            return GetLocalConstantRemovalInfo(target, context, count, itemIndex, items);
        }

        private static RewriterInfo GetModuleConstantRemovalInfo(
            VBAParser.ConstSubStmtContext target, VBAParser.ModuleDeclarationsElementContext element,
            int count, int itemIndex, IReadOnlyList<VBAParser.ConstSubStmtContext> items)
        {
            var startIndex = element.Start.TokenIndex;
            var parent = (VBAParser.ModuleDeclarationsContext)element.Parent;
            var elements = parent.moduleDeclarationsElement();

            if (count == 1)
            {
                var stopIndex = FindStopTokenIndex(elements, element, parent);
                return new RewriterInfo(startIndex, stopIndex);
            }
            return GetRewriterInfoForTargetRemovedFromListStmt(target.Start, itemIndex, items);
        }

        private static RewriterInfo GetLocalConstantRemovalInfo(VBAParser.ConstSubStmtContext target,
            VBAParser.ConstStmtContext constants,
            int count, int itemIndex, IReadOnlyList<VBAParser.ConstSubStmtContext> items)
        {
            var mainBlockStmt = (VBAParser.MainBlockStmtContext)constants.Parent;
            var startIndex = mainBlockStmt.Start.TokenIndex;
            var blockStmt = (VBAParser.BlockStmtContext)mainBlockStmt.Parent;
            var block = (VBAParser.BlockContext)blockStmt.Parent;
            var statements = block.blockStmt();

            if (count == 1)
            {
                var stopIndex = FindStopTokenIndex(statements, mainBlockStmt, block);
                return new RewriterInfo(startIndex, stopIndex);
            }
            return GetRewriterInfoForTargetRemovedFromListStmt(target.Start, itemIndex, items);
        }
    }
}