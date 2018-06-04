using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using System;
using static Rubberduck.Parsing.Grammar.VBAParser;

namespace Rubberduck.Inspections.QuickFixes
{
    public class AddStepOneQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public AddStepOneQuickFix(RubberduckParserState state)
            : base(typeof(StepIsNotSpecifiedInspection))
        {
            _state = state;
        }

        public override bool CanFixInProcedure => true;

        public override bool CanFixInModule => true;

        public override bool CanFixInProject => true;

        public override string Description(IInspectionResult result)
        {
            return Resources.Inspections.QuickFixes.AddStepOneQuickFix;
        }

        public override void Fix(IInspectionResult result)
        {
            IModuleRewriter rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);
            ForNextStmtContext context = result.Context as ForNextStmtContext;

            int toExpressionEnd = this.GetToExpressionEnd(context);
            rewriter.InsertAfter(toExpressionEnd, " Step 1");
        }

        private int GetToExpressionEnd(ForNextStmtContext context)
        {
            int toNodeIndex = context.TO().Symbol.TokenIndex;

            foreach(ExpressionContext expressionChild in context.expression())
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
