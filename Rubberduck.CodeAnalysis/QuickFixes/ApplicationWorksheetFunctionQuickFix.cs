using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class ApplicationWorksheetFunctionQuickFix : QuickFixBase
    {
        public ApplicationWorksheetFunctionQuickFix()
            : base(typeof(ApplicationWorksheetFunctionInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);
            rewriter.InsertBefore(result.Context.Start.TokenIndex, "WorksheetFunction.");
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.ApplicationWorksheetFunctionQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}
