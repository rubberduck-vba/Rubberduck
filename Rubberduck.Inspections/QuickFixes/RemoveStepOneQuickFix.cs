using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public class RemoveStepOneQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public RemoveStepOneQuickFix(RubberduckParserState state)
            : base(typeof(StepOneIsRedundantInspection))
        {
            _state = state;
        }

        public override bool CanFixInProcedure => true;

        public override bool CanFixInModule => true;

        public override bool CanFixInProject => true;

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RemoveStepOneQuickFix;

        public override void Fix(IInspectionResult result)
        {
            IModuleRewriter rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);
            var context = result.Context;
            rewriter.Remove(context);
        }
    }
}
