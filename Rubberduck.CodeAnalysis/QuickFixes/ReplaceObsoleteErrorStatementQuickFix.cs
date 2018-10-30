using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class ReplaceObsoleteErrorStatementQuickFix : QuickFixBase
    {
        public ReplaceObsoleteErrorStatementQuickFix()
            : base(typeof(ObsoleteErrorSyntaxInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);
            var context = (VBAParser.ErrorStmtContext) result.Context;

            rewriter.Replace(context.ERROR(), "Err.Raise");
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.ReplaceObsoleteErrorStatementQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}