using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Rewriter.RewriterInfo
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
            var count = items.Length;

            if (context.TryGetAncestor<VBAParser.ModuleDeclarationsElementContext>(out var element))
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
            if (count == 1)
            {
                return GetSeparateModuleVariableRemovalInfo(element);
            }
            return GetRewriterInfoForTargetRemovedFromListStmt(target.Start, itemIndex, items);
        }

        private static RewriterInfo GetSeparateModuleVariableRemovalInfo(VBAParser.ModuleDeclarationsElementContext element)
        {
            var startIndex = element.Start.TokenIndex;
            var stopIndex = FindStopTokenIndexForRemoval(element);
            return new RewriterInfo(startIndex, stopIndex);
        }

        private static RewriterInfo GetLocalVariableRemovalInfo(VBAParser.VariableSubStmtContext target,
            VBAParser.VariableListStmtContext variables,
            int count, int itemIndex, IReadOnlyList<VBAParser.VariableSubStmtContext> items)
        {
            if (count == 1)
            {
                return GetSeparateLocalVariableRemovalInfo(variables);
            }

            return GetRewriterInfoForTargetRemovedFromListStmt(target.Start, itemIndex, items);
        }

        private static RewriterInfo GetSeparateLocalVariableRemovalInfo(VBAParser.VariableListStmtContext variables)
        {
            var mainBlockStmt = variables.GetAncestor<VBAParser.MainBlockStmtContext>();
            var startIndex = mainBlockStmt.Start.TokenIndex;

            int stopIndex = FindStopTokenIndexForRemoval(mainBlockStmt);

            return new RewriterInfo(startIndex, stopIndex);
        }
    }
}