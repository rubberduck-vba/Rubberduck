using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class MakeSingleLineParameterQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public MakeSingleLineParameterQuickFix(RubberduckParserState state)
            : base(typeof(MultilineParameterInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);

            var parameter = result.Context.GetText()
                .Replace("_", "")
                .RemoveExtraSpacesLeavingIndentation();

            rewriter.Replace(result.Context, parameter);
        }

        public override string Description(IInspectionResult result) => InspectionsUI.MakeSingleLineParameterQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}
