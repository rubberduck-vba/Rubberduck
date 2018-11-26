using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class SpecifyExplicitPublicModifierQuickFix : QuickFixBase
    {
        public SpecifyExplicitPublicModifierQuickFix()
            : base(typeof(ImplicitPublicMemberInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.Target.QualifiedModuleName);
            rewriter.InsertBefore(result.Context.Start.TokenIndex, "Public ");
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.SpecifyExplicitPublicModifierQuickFix;

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}