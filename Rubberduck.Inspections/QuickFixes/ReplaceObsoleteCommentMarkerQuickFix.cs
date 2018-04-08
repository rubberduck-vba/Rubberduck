using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class ReplaceObsoleteCommentMarkerQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public ReplaceObsoleteCommentMarkerQuickFix(RubberduckParserState state)
            : base(typeof(ObsoleteCommentSyntaxInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);
            var context = (VBAParser.RemCommentContext) result.Context;

            rewriter.Replace(context.REM(), "'");
        }

        public override string Description(IInspectionResult result) => InspectionsUI.RemoveObsoleteStatementQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}