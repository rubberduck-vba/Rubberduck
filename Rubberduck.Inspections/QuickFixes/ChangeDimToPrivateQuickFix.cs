using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class ChangeDimToPrivateQuickFix : QuickFixBase, IQuickFix
    {
        private readonly RubberduckParserState _state;

        public ChangeDimToPrivateQuickFix(RubberduckParserState state, InspectionLocator inspectionLocator)
        {
            _state = state;
            RegisterInspections(inspectionLocator.GetInspection<ModuleScopeDimKeywordInspection>());
        }

        public void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);

            var context = (VBAParser.VariableStmtContext)result.Context.Parent.Parent;
            rewriter.Replace(context.DIM(), Tokens.Private);
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.ChangeDimToPrivateQuickFix;
        }

        public bool CanFixInProcedure { get; } = false;
        public bool CanFixInModule { get; } = true;
        public bool CanFixInProject { get; } = true;
    }
}