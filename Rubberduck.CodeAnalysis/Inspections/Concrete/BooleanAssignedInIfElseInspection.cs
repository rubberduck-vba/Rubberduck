using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies conditional assignments to mutually exclusive Boolean literal values in conditional branches.
    /// </summary>
    /// <why>
    /// The assignment could be made directly to the result of the conditional Boolean expression instead.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal value As Long)
    ///     Dim result As Boolean
    ///     If value > 10 Then
    ///         result = True
    ///     Else
    ///         result = False
    ///     End If
    ///     Debug.Print result
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal value As Long)
    ///     Dim result As Boolean
    ///     result = value > 10
    ///     Debug.Print result
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class BooleanAssignedInIfElseInspection : ParseTreeInspectionBase<VBAParser.IfStmtContext>
    {
        public BooleanAssignedInIfElseInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            ContextListener = new BooleanAssignedInIfElseListener();
        }
        
        protected override IInspectionListener<VBAParser.IfStmtContext> ContextListener { get; }

        protected override string ResultDescription(QualifiedContext<VBAParser.IfStmtContext> context)
        {
            var literalText = context.Context
                .block()
                .GetDescendent<VBAParser.LetStmtContext>()
                .lExpression()
                .GetText()
                .Trim();
            return string.Format(
                InspectionResults.BooleanAssignedInIfElseInspection, 
                literalText);
        }

        private class BooleanAssignedInIfElseListener : InspectionListenerBase<VBAParser.IfStmtContext>
        {
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

                SaveContext(context);
            }

            private static bool IsSingleBooleanAssignment(VBAParser.BlockContext block)
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
