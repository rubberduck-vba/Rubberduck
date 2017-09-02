using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class MakeSingleLineParameterQuickFix : QuickFixBase, IQuickFix
    {
        private readonly RubberduckParserState _state;

        public MakeSingleLineParameterQuickFix(RubberduckParserState state, InspectionLocator inspectionLocator)
        {
            _state = state;
            RegisterInspections(inspectionLocator.GetInspection<MultilineParameterInspection>());
        }

        public void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);

            var parameter = result.Context.GetText()
                .Replace("_", "")
                .RemoveExtraSpacesLeavingIndentation();

            rewriter.Replace(result.Context, parameter);
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.MakeSingleLineParameterQuickFix;
        }

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}
