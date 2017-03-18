using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.PostProcessing.RewriterInfo
{
    public class VariableRewriterInfoFinder : RewriterInfoFinderBase
    { 
        public override RewriterInfo GetRewriterInfo(ParserRuleContext context, Declaration target)
        {
            return GetRewriterInfo(target, context.Parent as VBAParser.VariableListStmtContext);
        }

        private static RewriterInfo GetRewriterInfo(Declaration target, VBAParser.VariableListStmtContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException("context", @"Context is null. Expecting a VBAParser.VariableListStmtContext instance.");
            }

            var items = context.variableSubStmt();
            var itemIndex = items.ToList().IndexOf((VBAParser.VariableSubStmtContext)target.Context);
            var count = items.Count;

            var element = context.Parent.Parent as VBAParser.ModuleDeclarationsElementContext;
            if (element != null)
            {
                return GetModuleVariableRemovalInfo(target, element, count, itemIndex, items);
            }

            if (context.Parent is VBAParser.VariableStmtContext)
            {
                return GetLocalVariableRemovalInfo(target, context, count, itemIndex, items);
            }

            return RewriterInfo.None;
        }

        private static RewriterInfo GetModuleVariableRemovalInfo(Declaration target,
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
            return GetRewriterInfoForTargetRemovedFromListStmt(target.Context.Start, itemIndex, items);
        }

        private static RewriterInfo GetLocalVariableRemovalInfo(Declaration target,
            VBAParser.VariableListStmtContext variables,
            int count, int itemIndex, IReadOnlyList<VBAParser.VariableSubStmtContext> items)
        {
            var blockStmt = (VBAParser.BlockStmtContext)variables.Parent.Parent;
            var startIndex = blockStmt.Start.TokenIndex;
            var parent = (VBAParser.BlockContext)blockStmt.Parent;
            var statements = parent.blockStmt();

            if (count == 1)
            {
                var stopIndex = FindStopTokenIndex(statements, blockStmt, parent);
                return new RewriterInfo(startIndex, stopIndex);
            }
            return GetRewriterInfoForTargetRemovedFromListStmt(target.Context.Start, itemIndex, items);
        }
    }
}