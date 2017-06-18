using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.PostProcessing.RewriterInfo
{
    public class VariableRewriterInfoFinder : RewriterInfoFinderBase
    {
        public override RewriterInfo GetRewriterInfo(ParserRuleContext context)
        {
            return GetRewriterInfo(context as VBAParser.VariableSubStmtContext, context.Parent as VBAParser.VariableListStmtContext);
        }

        private static RewriterInfo GetRewriterInfo(VBAParser.VariableSubStmtContext variable, VBAParser.VariableListStmtContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context), @"Context is null. Expecting a VBAParser.VariableListStmtContext instance.");
            }

            var items = context.variableSubStmt();
            var itemIndex = items.ToList().IndexOf(variable);
            var count = items.Count;

            var element = context.Parent.Parent as VBAParser.ModuleDeclarationsElementContext;
            if (element != null)
            {
                return GetModuleVariableRemovalInfo(variable, element, count, itemIndex, items);
            }

            if (context.Parent is VBAParser.VariableStmtContext)
            {
                return GetLocalVariableRemovalInfo(variable, context, count, itemIndex, items);
            }

            return RewriterInfo.None;
        }

        private static RewriterInfo GetModuleVariableRemovalInfo(VBAParser.VariableSubStmtContext target,
            VBAParser.ModuleDeclarationsElementContext element,
            int count, int itemIndex, IReadOnlyList<VBAParser.VariableSubStmtContext> items)
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

        private static RewriterInfo GetLocalVariableRemovalInfo(VBAParser.VariableSubStmtContext target,
            VBAParser.VariableListStmtContext variables,
            int count, int itemIndex, IReadOnlyList<VBAParser.VariableSubStmtContext> items)
        {
            var mainBlockStmt = (VBAParser.MainBlockStmtContext)variables.Parent.Parent;
            var startIndex = mainBlockStmt.Start.TokenIndex;
            var blockStmt = (VBAParser.BlockStmtContext)mainBlockStmt.Parent;
            var block = (VBAParser.BlockContext)blockStmt.Parent;
            var statements = block.blockStmt();

            if (count == 1)
            {
                var stopIndex = FindStopTokenIndex(statements, blockStmt, block);
                return new RewriterInfo(startIndex, stopIndex);
            }
            return GetRewriterInfoForTargetRemovedFromListStmt(target.Start, itemIndex, items);
        }
    }
}