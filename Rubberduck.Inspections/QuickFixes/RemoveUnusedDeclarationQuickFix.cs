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
    public sealed class RemoveUnusedDeclarationQuickFix : QuickFixBase, IQuickFix
    {
        private readonly RubberduckParserState _state;

        public RemoveUnusedDeclarationQuickFix(RubberduckParserState state, InspectionLocator inspectionLocator)
        {
            _state = state;
            RegisterInspections(inspectionLocator.GetInspection<ConstantNotUsedInspection>(),
                inspectionLocator.GetInspection<ProcedureNotUsedInspection>(),
                inspectionLocator.GetInspection<VariableNotUsedInspection>(),
                inspectionLocator.GetInspection<LineLabelNotUsedInspection>());
        }

        public void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.Target);
            rewriter.Remove(result.Target);
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.RemoveUnusedDeclarationQuickFix;
        }

        public bool CanFixInProcedure => false;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}