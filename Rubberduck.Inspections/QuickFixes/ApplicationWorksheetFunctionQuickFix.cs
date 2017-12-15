using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class ApplicationWorksheetFunctionQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public ApplicationWorksheetFunctionQuickFix(RubberduckParserState state)
            : base(typeof(ApplicationWorksheetFunctionInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);
            rewriter.InsertBefore(result.Context.Start.TokenIndex, "WorksheetFunction.");
        }

        public override string Description(IInspectionResult result) => InspectionsUI.ApplicationWorksheetFunctionQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}
