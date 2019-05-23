using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.CodeAnalysis.QuickFixes
{
    public sealed class ReplaceWhileWendWithDoWhileLoopQuickFix : QuickFixBase
    {
        public ReplaceWhileWendWithDoWhileLoopQuickFix()
            : base(typeof(ObsoleteWhileWendStatementInspection))
        { }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);
            var context = (VBAParser.WhileWendStmtContext)result.Context;

            rewriter.Replace(context.WHILE(), "Do While");
            rewriter.Replace(context.WEND(), "Loop");
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.ReplaceWhileWendWithDoWhileLoopQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}
