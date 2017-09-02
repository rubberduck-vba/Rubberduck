using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class PassParameterByReferenceQuickFix : QuickFixBase, IQuickFix
    {
        private readonly RubberduckParserState _state;

        public PassParameterByReferenceQuickFix(RubberduckParserState state, InspectionLocator inspectionLocator)
        {
            _state = state;
            RegisterInspections(inspectionLocator.GetInspection<AssignedByValParameterInspection>());
        }

        public void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.Target);

            var token = ((VBAParser.ArgContext)result.Target.Context).BYVAL().Symbol;
            rewriter.Replace(token, Tokens.ByRef);
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.PassParameterByReferenceQuickFix;
        }

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}