using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class ReplaceObsoleteErrorStatementQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public ReplaceObsoleteErrorStatementQuickFix(RubberduckParserState state)
            : base(typeof(ObsoleteErrorSyntaxInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);
            var context = (VBAParser.ErrorStmtContext) result.Context;

            rewriter.Replace(context.ERROR(), "Err.Raise");
        }

        public override string Description(IInspectionResult result) => InspectionsUI.ReplaceObsoleteErrorStatementQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}