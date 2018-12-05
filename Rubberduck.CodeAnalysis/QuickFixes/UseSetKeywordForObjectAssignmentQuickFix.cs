using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
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
            var letStmt = result.Context.GetAncestor<VBAParser.LetStmtContext>();
            var letToken = letStmt.LET();
            if (letToken != null)
            {
                rewriter.Replace(letToken, "Set");
            }
            else
            {
                rewriter.InsertBefore(letStmt.Start.TokenIndex, "Set ");
            }
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.SetObjectVariableQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}