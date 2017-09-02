using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class ApplicationWorksheetFunctionQuickFix : QuickFixBase, IQuickFix
    {
        private readonly RubberduckParserState _state;

        public ApplicationWorksheetFunctionQuickFix(RubberduckParserState state, InspectionLocator inspectionLocator)
        {
            _state = state;
            RegisterInspections(inspectionLocator.GetInspection<ApplicationWorksheetFunctionInspection>());
        }

        public void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);
            rewriter.InsertBefore(result.Context.Start.TokenIndex, "WorksheetFunction.");
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.ApplicationWorksheetFunctionQuickFix;
        }

        public bool CanFixInProcedure { get; } = true;
        public bool CanFixInModule { get; } = true;
        public bool CanFixInProject { get; } = true;
    }
}
