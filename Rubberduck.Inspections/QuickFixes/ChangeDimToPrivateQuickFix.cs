using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class ChangeDimToPrivateQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public ChangeDimToPrivateQuickFix(RubberduckParserState state)
            : base(typeof(ModuleScopeDimKeywordInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);

            var context = (VBAParser.VariableStmtContext)result.Context.Parent.Parent;
            rewriter.Replace(context.DIM(), Tokens.Private);
        }

        public override string Description(IInspectionResult result) => InspectionsUI.ChangeDimToPrivateQuickFix;

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}