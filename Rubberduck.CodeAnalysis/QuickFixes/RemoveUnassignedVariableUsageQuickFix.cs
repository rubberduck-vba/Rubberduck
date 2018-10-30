using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RemoveUnassignedVariableUsageQuickFix : QuickFixBase
    {
        public RemoveUnassignedVariableUsageQuickFix()
            : base(typeof(UnassignedVariableUsageInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);

            var assignmentContext = result.Context.GetAncestor<VBAParser.LetStmtContext>() ??
                                                  (ParserRuleContext)result.Context.GetAncestor<VBAParser.CallStmtContext>();

            rewriter.Remove(assignmentContext);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RemoveUnassignedVariableUsageQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}