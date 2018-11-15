using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class ReplaceEmptyStringLiteralStatementQuickFix : QuickFixBase
    {
        public ReplaceEmptyStringLiteralStatementQuickFix()
            : base(typeof(EmptyStringLiteralInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);
            rewriter.Replace(result.Context, "vbNullString");
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.EmptyStringLiteralInspectionQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}