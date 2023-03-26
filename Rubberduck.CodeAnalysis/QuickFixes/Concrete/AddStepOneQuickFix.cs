using System;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Rewriter;
using static Rubberduck.Parsing.Grammar.VBAParser;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Adds an explicit Step specifier to a For loop instruction.
    /// </summary>
    /// <inspections>
    /// <inspection name="StepIsNotSpecifiedInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim i As Long
    ///     For i = 1 To 10
    ///         Debug.Print i
    ///     Next
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim i As Long
    ///     For i = 1 To 10 Step 1
    ///         Debug.Print i
    ///     Next
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class AddStepOneQuickFix : QuickFixBase
    {
        public AddStepOneQuickFix()
            : base(typeof(StepIsNotSpecifiedInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);
            var context = result.Context as ForNextStmtContext;

            var toExpressionEnd = GetToExpressionEnd(context);
            rewriter.InsertAfter(toExpressionEnd, " Step 1");
        }

        private static int GetToExpressionEnd(ForNextStmtContext context)
        {
            var toNodeIndex = context.TO().Symbol.TokenIndex;

            foreach(var expressionChild in context.expression())
            {
                if (expressionChild.Stop.TokenIndex > toNodeIndex)
                {
                    return expressionChild.Stop.TokenIndex;
                }
            }

            throw new InvalidOperationException();
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.AddStepOneQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}
