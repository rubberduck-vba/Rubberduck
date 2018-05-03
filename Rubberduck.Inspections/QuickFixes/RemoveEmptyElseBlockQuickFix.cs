using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    class RemoveEmptyElseBlockQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public RemoveEmptyElseBlockQuickFix(RubberduckParserState state)
            : base(typeof(EmptyElseBlockInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);

            //dynamic used since it's not known at run-time
            UpdateContext((dynamic)result.Context, rewriter);
        }

        private void UpdateContext(VBAParser.ElseBlockContext context, IModuleRewriter rewriter)
        {
            var elseBlock = context.block();

            if (elseBlock.ChildCount == 0 )
            {
                rewriter.Remove(context);
            }
        }
        
        public override string Description(IInspectionResult result) => InspectionsUI.RemoveEmptyElseBlockQuickFix;

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => false;
    }
}
