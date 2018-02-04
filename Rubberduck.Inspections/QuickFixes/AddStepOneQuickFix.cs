using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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

        public override bool CanFixInProcedure => false;

        public override bool CanFixInModule => false;

        public override bool CanFixInProject => false;

        public override string Description(IInspectionResult result)
        {
            return InspectionsUI.AddStepOneQuickFix;
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
