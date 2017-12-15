using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class BooleanAssignedInIfElseInspection : ParseTreeInspectionBase
    {
        public BooleanAssignedInIfElseInspection(RubberduckParserState state)
            : base(state) { }
        
        public override CodeInspectionType InspectionType => CodeInspectionType.MaintainabilityAndReadabilityIssues;

        public override IInspectionListener Listener { get; } =
            new BooleanAssignedInIfElseListener();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line))
                .Select(result => new QualifiedContextInspectionResult(this,
                                                       string.Format(InspectionsUI.BooleanAssignedInIfElseInspectionResultFormat,
                                                            ParserRuleContextHelper.GetDescendent<VBAParser.LetStmtContext>(((VBAParser.IfStmtContext)result.Context).block()).lExpression().GetText().Trim()),
                                                       result));
        }

        public class BooleanAssignedInIfElseListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;
            
            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void ExitIfStmt(VBAParser.IfStmtContext context)
            {
                if (context.elseIfBlock() != null && context.elseIfBlock().Any())
                {
                    return;
                }

                if (context.elseBlock() == null)
                {
                    return;
                }

                if (!IsSingleBooleanAssignment(context.block()) ||
                    !IsSingleBooleanAssignment(context.elseBlock().block()))
                {
                    return;
                }

                // make sure the assignments are the opposite

                if (!(ParserRuleContextHelper.GetDescendent<VBAParser.BooleanLiteralIdentifierContext>(context.block()).GetText() == Tokens.True ^
                      ParserRuleContextHelper.GetDescendent<VBAParser.BooleanLiteralIdentifierContext>(context.elseBlock().block()).GetText() == Tokens.True))
                {
                    return;
                }

                if (ParserRuleContextHelper.GetDescendent<VBAParser.LetStmtContext>(context.block()).lExpression().GetText().ToLowerInvariant() !=
                    ParserRuleContextHelper.GetDescendent<VBAParser.LetStmtContext>(context.elseBlock().block()).lExpression().GetText().ToLowerInvariant())
                {
                    return;
                }

                _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
            }

            private bool IsSingleBooleanAssignment(VBAParser.BlockContext block)
            {
                if (block.ChildCount != 2)
                {
                    return false;
                }

                var mainBlockStmtContext =
                    ParserRuleContextHelper.GetDescendent<VBAParser.MainBlockStmtContext>(block);

                return mainBlockStmtContext.children.FirstOrDefault() is VBAParser.LetStmtContext letStmt &&
                       letStmt.expression() is VBAParser.LiteralExprContext literal &&
                       ParserRuleContextHelper.GetDescendent<VBAParser.BooleanLiteralIdentifierContext>(literal) != null;
            }
        }
    }
}
