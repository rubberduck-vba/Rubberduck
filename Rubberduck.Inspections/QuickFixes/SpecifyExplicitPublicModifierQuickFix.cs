using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class SpecifyExplicitPublicModifierQuickFix : QuickFixBase, IQuickFix
    {
        private readonly RubberduckParserState _state;

        public SpecifyExplicitPublicModifierQuickFix(RubberduckParserState state, InspectionLocator inspectionLocator)
        {
            _state = state;
            RegisterInspections(inspectionLocator.GetInspection<ImplicitPublicMemberInspection>());
        }

        public void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.Target);
            rewriter.InsertBefore(result.Context.Start.TokenIndex, "Public ");
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.SpecifyExplicitPublicModifierQuickFix;
        }

        public bool CanFixInProcedure => false;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}