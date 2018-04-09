using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class SpecifyExplicitPublicModifierQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public SpecifyExplicitPublicModifierQuickFix(RubberduckParserState state)
            : base(typeof(ImplicitPublicMemberInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.Target);
            rewriter.InsertBefore(result.Context.Start.TokenIndex, "Public ");
        }

        public override string Description(IInspectionResult result) => InspectionsUI.SpecifyExplicitPublicModifierQuickFix;

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}