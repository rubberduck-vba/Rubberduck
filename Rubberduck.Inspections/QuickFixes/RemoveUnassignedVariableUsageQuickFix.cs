using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RemoveUnassignedVariableUsageQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public RemoveUnassignedVariableUsageQuickFix(RubberduckParserState state)
            : base(typeof(UnassignedVariableUsageInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);

            var assignmentContext = result.Context.GetAncestor<VBAParser.LetStmtContext>() ??
                                                  (ParserRuleContext)result.Context.GetAncestor<VBAParser.CallStmtContext>();

            rewriter.Remove(assignmentContext);
        }

        public override string Description(IInspectionResult result) => InspectionsUI.RemoveUnassignedVariableUsageQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}