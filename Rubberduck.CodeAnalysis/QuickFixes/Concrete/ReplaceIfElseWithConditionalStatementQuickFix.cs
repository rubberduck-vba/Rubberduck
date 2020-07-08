using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Simplifies conditional Boolean literal assignments with a direct assignment to the conditional expression.
    /// </summary>
    /// <inspections>
    /// <inspection name="BooleanAssignedInIfElseInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// 
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
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething(ByVal value As Long)
    ///     Dim result As Boolean
    ///     result = value > 10
    ///     Debug.Print result
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class ReplaceIfElseWithConditionalStatementQuickFix : QuickFixBase
    {
        public ReplaceIfElseWithConditionalStatementQuickFix()
            : base(typeof(BooleanAssignedInIfElseInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var ifContext = (VBAParser.IfStmtContext) result.Context;
            var letStmt = ifContext.block().GetDescendent<VBAParser.LetStmtContext>();

            var conditional = ifContext.booleanExpression().GetText();

            if (letStmt.expression().GetText() == Tokens.False)
            {
                conditional = $"Not ({conditional})";
            }

            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);
            rewriter.Replace(result.Context, $"{letStmt.lExpression().GetText()} = {conditional}");
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.ReplaceIfElseWithConditionalStatementQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}
