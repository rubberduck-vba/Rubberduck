using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    /// <summary>
    /// A code inspection quickfix that removes an unused identifier declaration.
    /// </summary>
    public sealed class RemoveUnusedDeclarationQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public RemoveUnusedDeclarationQuickFix(RubberduckParserState state)
            : base(typeof(ConstantNotUsedInspection), typeof(ProcedureNotUsedInspection), typeof(VariableNotUsedInspection), typeof(LineLabelNotUsedInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.Target);
            rewriter.Remove(result.Target);
        }

        public override string Description(IInspectionResult result) => InspectionsUI.RemoveUnusedDeclarationQuickFix;

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}