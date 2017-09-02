using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class ReplaceGlobalModifierQuickFix : QuickFixBase, IQuickFix
    {
        private readonly RubberduckParserState _state;

        public ReplaceGlobalModifierQuickFix(RubberduckParserState state, InspectionLocator inspectionLocator)
        {
            _state = state;
            RegisterInspections(inspectionLocator.GetInspection<ObsoleteGlobalInspection>());
        }

        public void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.Target);
            rewriter.Replace(ParserRuleContextHelper.GetDescendent<VBAParser.VisibilityContext>(result.Context.Parent.Parent), Tokens.Public);
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.ObsoleteGlobalInspectionQuickFix;
        }

        public bool CanFixInProcedure => false;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}