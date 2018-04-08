using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RemoveExplicitLetStatementQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public RemoveExplicitLetStatementQuickFix(RubberduckParserState state)
            : base(typeof(ObsoleteLetStatementInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);

            var context = (VBAParser.LetStmtContext) result.Context;
            rewriter.Remove(context.LET());
            rewriter.Remove(context.whiteSpace().First());
        }

        public override string Description(IInspectionResult result) => InspectionsUI.RemoveObsoleteStatementQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}