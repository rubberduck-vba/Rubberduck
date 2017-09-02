using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class ReplaceEmptyStringLiteralStatementQuickFix : QuickFixBase, IQuickFix
    {
        private readonly RubberduckParserState _state;

        public ReplaceEmptyStringLiteralStatementQuickFix(RubberduckParserState state, InspectionLocator inspectionLocator)
        {
            _state = state;
            RegisterInspections(inspectionLocator.GetInspection<EmptyStringLiteralInspection>());
        }

        public void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);
            rewriter.Replace(result.Context, "vbNullString");
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.EmptyStringLiteralInspectionQuickFix;
        }

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}