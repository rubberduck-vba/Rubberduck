using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class PassParameterByReferenceQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public PassParameterByReferenceQuickFix(RubberduckParserState state)
            : base(typeof(AssignedByValParameterInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.Target);

            var token = ((VBAParser.ArgContext)result.Target.Context).BYVAL().Symbol;
            rewriter.Replace(token, Tokens.ByRef);
        }

        public override string Description(IInspectionResult result) => InspectionsUI.PassParameterByReferenceQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}