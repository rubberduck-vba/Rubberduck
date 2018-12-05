using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RemoveExplicitCallStatementQuickFix : QuickFixBase
    {
        public RemoveExplicitCallStatementQuickFix()
            : base(typeof(ObsoleteCallStatementInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);

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

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RemoveObsoleteStatementQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}
