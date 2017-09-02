using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class ReplaceObsoleteCommentMarkerQuickFix : QuickFixBase, IQuickFix
    {
        private readonly RubberduckParserState _state;

        public ReplaceObsoleteCommentMarkerQuickFix(RubberduckParserState state, InspectionLocator inspectionLocator)
        {
            _state = state;
            RegisterInspections(inspectionLocator.GetInspection<ObsoleteCommentSyntaxInspection>());
        }

        public void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);
            var context = (VBAParser.RemCommentContext) result.Context;

            rewriter.Replace(context.REM(), "'");
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.RemoveObsoleteStatementQuickFix;
        }

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}