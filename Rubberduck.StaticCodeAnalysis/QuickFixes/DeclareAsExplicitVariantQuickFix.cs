using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class DeclareAsExplicitVariantQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;
        
        public DeclareAsExplicitVariantQuickFix(RubberduckParserState state)
            : base(typeof(VariableTypeNotDeclaredInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.Target);

            ParserRuleContext identifierNode =
                result.Context is VBAParser.VariableSubStmtContext || result.Context is VBAParser.ConstSubStmtContext
                ? result.Context.children[0]
                : ((dynamic) result.Context).unrestrictedIdentifier();
            rewriter.InsertAfter(identifierNode.Stop.TokenIndex, " As Variant");
        }

        public override string Description(IInspectionResult result) => InspectionsUI.DeclareAsExplicitVariantQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}