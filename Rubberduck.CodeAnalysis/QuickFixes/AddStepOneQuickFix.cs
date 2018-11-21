using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using System;
using static Rubberduck.Parsing.Grammar.VBAParser;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class AddStepOneQuickFix : QuickFixBase
    {
        public AddStepOneQuickFix()
            : base(typeof(StepIsNotSpecifiedInspection))
        {}

        public override bool CanFixInProcedure => true;

        public override bool CanFixInModule => true;

        public override bool CanFixInProject => true;

        public override string Description(IInspectionResult result)
        {
            return Resources.Inspections.QuickFixes.AddStepOneQuickFix;
        }

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
    }
}
