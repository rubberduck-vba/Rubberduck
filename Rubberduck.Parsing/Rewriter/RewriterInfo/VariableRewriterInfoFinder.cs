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
            var count = items.Count;

            if (context.Parent.Parent is VBAParser.ModuleDeclarationsElementContext element)
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
            if (count == 1)
            {
                int stopIndex = variables.Stop.TokenIndex + 1; // also remove trailing newlines?
                
                var containingBlock = (VBAParser.BlockContext)mainBlockStmt.Parent.Parent;
                var blockStmtIndex = containingBlock.children.IndexOf(mainBlockStmt.Parent);
                // a few things can happen here
                if (blockStmtIndex == containingBlock.ChildCount)
                {
                    // well we're lucky?
                    stopIndex = containingBlock.Stop.TokenIndex;
                }
                else if (containingBlock.GetChild(blockStmtIndex + 1) is VBAParser.EndOfStatementContext eos)
                {
                    // remove until the end of the EOS or continue to the start of the following
                    if (blockStmtIndex + 2 >= containingBlock.ChildCount)
                    {
                        // since EOS includes the following comment, we need to do weird shit
                        // eos cannot be EOF, since we're on a local var
                        var eol = eos.endOfLine(0);
                        if (eol.commentOrAnnotation() != null)
                        {
                            stopIndex = eol.commentOrAnnotation().Start.TokenIndex - 1;
                        }
                        else
                        {
                            stopIndex = eol.Stop.TokenIndex;
                        }
                    } 
                    else
                    {
                        stopIndex = containingBlock.GetChild<ParserRuleContext>(blockStmtIndex + 2).Start.TokenIndex - 1;
                    }
                }
                // should be assigned by now, I guess
                /*
                var variableListStmtIndex = mainBlockStmt.children.IndexOf(variables.Parent);
                var nextStmt = mainBlockStmt.GetChild(Math.Min(variableListStmtIndex + 1, mainBlockStmt.children.Count - 1)) as ParserRuleContext;
              
                if (nextStmt == null)
                {
                    // it makes no sense that the next statment would not be a ParserRuleContext
                    return RewriterInfo.None;
                }
                else if (nextStmt == variables)
                {
                    stopIndex = variables.Stop.TokenIndex;
                }
                else
                {
                    stopIndex = nextStmt.Start.TokenIndex - 1;
                }
                */
                // we know the target is all we want to remove, so we'll use that
                // we also need to remove all tokens until the next useful statement
                //var nextStatement = mainBlockStmt.Parent.GetChild(itemIndex + 1);
                //int tokenIndex;
                //if (nextStatement is ParserRuleContext ctx)
                //{
                //    tokenIndex = ctx.Start.TokenIndex - 1;
                //}
                //else
                //{
                //    // it makes no sense that the next statment would not be a ParserRuleContext
                //    return RewriterInfo.None;
                //}

                return new RewriterInfo(startIndex, stopIndex);
            }

            var blockStmt = (VBAParser.BlockStmtContext)mainBlockStmt.Parent;
            var block = (VBAParser.BlockContext)blockStmt.Parent;
            var statements = block.blockStmt();
            return GetRewriterInfoForTargetRemovedFromListStmt(target.Start, itemIndex, items);
        }
    }
}