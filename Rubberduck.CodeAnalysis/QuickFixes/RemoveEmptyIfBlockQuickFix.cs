﻿using System;
using System.Diagnostics;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RemoveEmptyIfBlockQuickFix : QuickFixBase
    {
        public RemoveEmptyIfBlockQuickFix()
            : base(typeof(EmptyIfBlockInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);

            switch (result.Context)
            {
                case VBAParser.IfStmtContext ifContext:
                    UpdateContext(ifContext, rewriter);
                    break;
                case VBAParser.IfWithEmptyThenContext ifWithEmtyThenContext:
                    UpdateContext(ifWithEmtyThenContext, rewriter);
                    break;
                case VBAParser.ElseIfBlockContext elseIfBlockContext:
                    UpdateContext(elseIfBlockContext, rewriter);
                    break;
                default:
                    throw new NotSupportedException(result.Context.GetType().ToString());
            }
        }

        private void UpdateContext(VBAParser.IfStmtContext context, IModuleRewriter rewriter)
        {
            var elseBlock = context.elseBlock();
            var elseIfBlock = context.elseIfBlock().FirstOrDefault();

            if (BlockHasDeclaration(context.block()))
            {
                rewriter.InsertBefore(context.Start.TokenIndex, AdjustedBlockText(context.block()));
            }

            if (elseIfBlock != null)
            {
                rewriter.RemoveRange(context.IF().Symbol.TokenIndex, context.block()?.Stop.TokenIndex ?? context.endOfStatement().Stop.TokenIndex);
                rewriter.Replace(elseIfBlock.ELSEIF(), "If");
            }
            else if (elseBlock != null)
            {
                if (!string.IsNullOrEmpty(context.block()?.GetText()))
                {
                    rewriter.RemoveRange(context.block().Start.TokenIndex, elseBlock.ELSE().Symbol.TokenIndex);
                }
                else
                {
                    rewriter.Remove(elseBlock.ELSE());
                }

                Debug.Assert(context.booleanExpression().children.Count == 1);
                UpdateConditionDependingOnType((ParserRuleContext)context.booleanExpression().children[0], rewriter);
            }
            else
            {
                rewriter.Remove(context);
            }
        }

        private void UpdateContext(VBAParser.IfWithEmptyThenContext context, IModuleRewriter rewriter)
        {
            var elseClause = context.singleLineElseClause();
            if (context.singleLineElseClause().whiteSpace() != null)
            {
                rewriter.RemoveRange(elseClause.ELSE().Symbol.TokenIndex, elseClause.whiteSpace().Stop.TokenIndex);
            }
            else
            {
                rewriter.Remove(elseClause.ELSE());
            }

            Debug.Assert(context.booleanExpression().children.Count == 1);
            UpdateConditionDependingOnType((ParserRuleContext)context.booleanExpression().children[0], rewriter);
        }

        private void UpdateContext(VBAParser.ElseIfBlockContext context, IModuleRewriter rewriter)
        {
            if (BlockHasDeclaration(context.block()))
            {
                rewriter.InsertBefore(((VBAParser.IfStmtContext)context.Parent).Start.TokenIndex, AdjustedBlockText(context.block()));
            }

            rewriter.Remove(context);
        }

        private void UpdateConditionDependingOnType(ParserRuleContext context, IModuleRewriter rewriter)
        {
            switch (context)
            {
                case VBAParser.RelationalOpContext condition:
                    UpdateCondition(condition, rewriter);
                    break;
                case VBAParser.LogicalNotOpContext condition:
                    UpdateCondition(condition, rewriter);
                    break;
                case VBAParser.LogicalAndOpContext condition:
                    UpdateCondition(condition, rewriter);
                    break;
                case VBAParser.LogicalOrOpContext condition:
                    UpdateCondition(condition, rewriter);
                    break;
                default:
                    UpdateCondition(context, rewriter);
                    break;
            }
        }

        private void UpdateCondition(VBAParser.RelationalOpContext condition, IModuleRewriter rewriter)
        {
            if (condition.EQ() != null)
            {
                rewriter.Replace(condition.EQ(), "<>");
            }
            if (condition.NEQ() != null)
            {
                rewriter.Replace(condition.NEQ(), "=");
            }
            if (condition.LT() != null)
            {
                rewriter.Replace(condition.LT(), ">=");
            }
            if (condition.GT() != null)
            {
                rewriter.Replace(condition.GT(), "<=");
            }
            if (condition.LEQ() != null)
            {
                rewriter.Replace(condition.LEQ(), ">");
            }
            if (condition.GEQ() != null)
            {
                rewriter.Replace(condition.GEQ(), "<");
            }
            if (condition.IS() != null || condition.LIKE() != null)
            {
                UpdateCondition((ParserRuleContext)condition, rewriter);
            }
        }

        private void UpdateCondition(VBAParser.LogicalNotOpContext condition, IModuleRewriter rewriter)
        {
            if (condition.whiteSpace() != null)
            {
                rewriter.RemoveRange(condition.NOT().Symbol.TokenIndex, condition.whiteSpace().Stop.TokenIndex);
            }
            else
            {
                rewriter.Remove(condition.NOT());
            }
        }

        private void UpdateCondition(VBAParser.LogicalAndOpContext condition, IModuleRewriter rewriter)
        {
            rewriter.Replace(condition.AND(), "Or");
        }

        private void UpdateCondition(VBAParser.LogicalOrOpContext condition, IModuleRewriter rewriter)
        {
            rewriter.Replace(condition.OR(), "And");
        }

        private void UpdateCondition(ParserRuleContext condition, IModuleRewriter rewriter)
        {
            if (condition.GetText().Contains(' '))
            {
                rewriter.InsertBefore(condition.Start.TokenIndex, "Not (");
                rewriter.InsertAfter(condition.Stop.TokenIndex, ")");
            }
            else
            {
                rewriter.InsertBefore(condition.Start.TokenIndex, "Not ");
            }
        }

        private string AdjustedBlockText(VBAParser.BlockContext blockContext)
        {
            var blockText = blockContext.GetText();
            if (FirstBlockStmntHasLabel(blockContext))
            {
                blockText = Environment.NewLine + blockText;
            }

            return blockText;
        }

        private bool BlockHasDeclaration(VBAParser.BlockContext block)
            => block.blockStmt()?.Any() ?? false;

        private bool FirstBlockStmntHasLabel(VBAParser.BlockContext block)
            => block.blockStmt()?.FirstOrDefault()?.statementLabelDefinition() != null;

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RemoveEmptyIfBlockQuickFix;

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => false;
    }
}
