using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RemoveExplicitCallStatmentQuickFix : QuickFixBase, IQuickFix
    {
        private readonly RubberduckParserState _state;

        public RemoveExplicitCallStatmentQuickFix(RubberduckParserState state, InspectionLocator inspectionLocator)
        {
            _state = state;
            RegisterInspections(inspectionLocator.GetInspection<ObsoleteCallStatementInspection>());
        }

        public void Fix(IInspectionResult result)
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

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.RemoveObsoleteStatementQuickFix;
        }

        public bool CanFixInProcedure => true;

        public bool CanFixInModule => true;

        public bool CanFixInProject => true;
    }
}
