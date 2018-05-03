using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RemoveUnassignedIdentifierQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public RemoveUnassignedIdentifierQuickFix(RubberduckParserState state)
            : base(typeof(VariableNotAssignedInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.Target);
            rewriter.Remove(result.Target);
        }

        public override string Description(IInspectionResult result) => InspectionsUI.RemoveUnassignedIdentifierQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}