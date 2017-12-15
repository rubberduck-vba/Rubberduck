using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RemoveStopKeywordQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public RemoveStopKeywordQuickFix(RubberduckParserState state)
            : base(typeof(StopKeywordInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);
            rewriter.Remove(result.Context);
        }

        public override string Description(IInspectionResult result) => InspectionsUI.RemoveStopKeywordQuickFix;

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => false;
    }
}