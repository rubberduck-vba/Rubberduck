using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RemoveEmptyElseBlockQuickFix : QuickFixBase
    {
        public RemoveEmptyElseBlockQuickFix()
            : base(typeof(EmptyElseBlockInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);

            UpdateContext((VBAParser.ElseBlockContext)result.Context, rewriter);
        }

        private void UpdateContext(VBAParser.ElseBlockContext context, IModuleRewriter rewriter)
        {
            var elseBlock = context.block();

            if (elseBlock.ChildCount == 0 )
            {
                rewriter.Remove(context);
            }
        }
        
        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RemoveEmptyElseBlockQuickFix;

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => false;
    }
}
