using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RemoveExplicitCallStatmentQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public RemoveExplicitCallStatmentQuickFix(RubberduckParserState state)
            : base(typeof(ObsoleteCallStatementInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);

            var context = (VBAParser.CallStmtContext)result.Context;
            rewriter.Remove(context.CALL());
            rewriter.Remove(context.whiteSpace());

            // The CALL statement only has arguments if it's an index expression.
            if (context.lExpression() is VBAParser.IndexExprContext indexExpr)
            {
                rewriter.Replace(indexExpr.LPAREN(), " ");
                rewriter.Remove(indexExpr.RPAREN());
            }
        }

        public override string Description(IInspectionResult result) => InspectionsUI.RemoveObsoleteStatementQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}
