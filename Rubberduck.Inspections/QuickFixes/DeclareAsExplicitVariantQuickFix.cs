using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class DeclareAsExplicitVariantQuickFix : QuickFixBase, IQuickFix
    {
        private readonly RubberduckParserState _state;
        
        public DeclareAsExplicitVariantQuickFix(RubberduckParserState state, InspectionLocator inspectionLocator)
        {
            _state = state;
            RegisterInspections(inspectionLocator.GetInspection<VariableTypeNotDeclaredInspection>());
        }

        public void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.Target);

            ParserRuleContext identifierNode =
                result.Context is VBAParser.VariableSubStmtContext || result.Context is VBAParser.ConstSubStmtContext
                ? result.Context.children[0]
                : ((dynamic) result.Context).unrestrictedIdentifier();
            rewriter.InsertAfter(identifierNode.Stop.TokenIndex, " As Variant");
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.DeclareAsExplicitVariantQuickFix;
        }

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}