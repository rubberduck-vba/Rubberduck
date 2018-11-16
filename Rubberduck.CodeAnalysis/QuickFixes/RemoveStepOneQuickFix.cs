using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RemoveStepOneQuickFix : QuickFixBase
    {
        public RemoveStepOneQuickFix()
            : base(typeof(StepOneIsRedundantInspection))
        {}

        public override bool CanFixInProcedure => true;

        public override bool CanFixInModule => true;

        public override bool CanFixInProject => true;

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RemoveStepOneQuickFix;

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);
            var context = result.Context;
            rewriter.Remove(context);
        }
    }
}
