using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.PostProcessing;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    internal sealed class RemoveEmptyIfBlockQuickFix : IQuickFix
    {
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type> { typeof(EmptyIfBlockInspection) };
        private readonly RubberduckParserState _state;

        public RemoveEmptyIfBlockQuickFix(RubberduckParserState state)
        {
            _state = state;
        }

        public IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);

            UpdateContext((dynamic)result.Context, rewriter);
        }

        private void UpdateContext(VBAParser.IfStmtContext context, IModuleRewriter rewriter)
        {
            var elseBlock = context.elseBlock();

            if (elseBlock == null)
            {
                var elseIfBlock = context.elseIfBlock().FirstOrDefault();
                if (elseIfBlock != null)
                {
                    rewriter.RemoveRange(context.IF().Symbol.TokenIndex, context.endOfStatement().Stop.TokenIndex);
                    rewriter.Replace(elseIfBlock.ELSEIF(), "If");
                }
                else
                {
                    rewriter.Remove(context);
                }
            }
            else
            {
                rewriter.Remove(elseBlock.ELSE());

                Debug.Assert(context.booleanExpression().children.Count == 1);
                UpdateCondition((dynamic)context.booleanExpression().children[0], rewriter);
            }
        }

        private void UpdateContext(VBAParser.ElseIfBlockContext context, IModuleRewriter rewriter)
        {
            rewriter.Remove(context);
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
            rewriter.InsertBefore(condition.Start.TokenIndex, "Not (");
            rewriter.InsertAfter(condition.Stop.TokenIndex, ")");
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.RemoveEmptyIfBlockQuickFix;
        }

        public bool CanFixInProcedure { get; } = false;
        public bool CanFixInModule { get; } = true;
        public bool CanFixInProject { get; } = true;
    }
}
