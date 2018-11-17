using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class UseSetKeywordForObjectAssignmentQuickFix : QuickFixBase
    {
        public UseSetKeywordForObjectAssignmentQuickFix()
            : base(typeof(ObjectVariableNotSetInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);
            rewriter.InsertBefore(result.Context.Start.TokenIndex, "Set ");
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.SetObjectVariableQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}