using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Identifies redundant Boolean expressions in conditionals.
    /// </summary>
    /// <why>
    /// A Boolean expression never needs to be compared to a Boolean literal in a conditional expression.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Boolean)
    ///     If foo = True Then ' foo is known to already be a Boolean value.
    ///         ' ...
    ///     End If
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Boolean)
    ///     If foo Then
    ///         ' ...
    ///     End If
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class BooleanAssignedInIfElseInspection : ParseTreeInspectionBase
    {
        public BooleanAssignedInIfElseInspection(RubberduckParserState state)
            : base(state) { }
        
        public override IInspectionListener Listener { get; } =
            new BooleanAssignedInIfElseListener();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts
                .Select(result => new QualifiedContextInspectionResult(this,
                                                       string.Format(InspectionResults.BooleanAssignedInIfElseInspection,
                                                            (((VBAParser.IfStmtContext)result.Context).block().GetDescendent<VBAParser.LetStmtContext>()).lExpression().GetText().Trim()),
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

                if (!(context.block().GetDescendent<VBAParser.BooleanLiteralIdentifierContext>().GetText() == Tokens.True ^
                      context.elseBlock().block().GetDescendent<VBAParser.BooleanLiteralIdentifierContext>().GetText() == Tokens.True))
                {
                    return;
                }

                if (context.block().GetDescendent<VBAParser.LetStmtContext>().lExpression().GetText().ToLowerInvariant() !=
                    context.elseBlock().block().GetDescendent<VBAParser.LetStmtContext>().lExpression().GetText().ToLowerInvariant())
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

                var mainBlockStmtContext = block.GetDescendent<VBAParser.MainBlockStmtContext>();

                return mainBlockStmtContext.children.FirstOrDefault() is VBAParser.LetStmtContext letStmt &&
                       letStmt.expression() is VBAParser.LiteralExprContext literal &&
                       literal.GetDescendent<VBAParser.BooleanLiteralIdentifierContext>() != null;
            }
        }
    }
}
