using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RemoveUnassignedIdentifierQuickFix : QuickFixBase
    {
        public RemoveUnassignedIdentifierQuickFix()
            : base(typeof(VariableNotAssignedInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.Target.QualifiedModuleName);
            rewriter.Remove(result.Target);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RemoveUnassignedIdentifierQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}