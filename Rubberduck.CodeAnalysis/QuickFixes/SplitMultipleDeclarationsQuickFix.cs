using System.Text;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class SplitMultipleDeclarationsQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public SplitMultipleDeclarationsQuickFix(RubberduckParserState state)
            : base(typeof(MultipleDeclarationsInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result)
        {
            dynamic context = result.Context is VBAParser.ConstStmtContext
                ? result.Context
                : result.Context.Parent;

            var declarationsText = GetDeclarationsText(context);

            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);
            rewriter.Replace(context, declarationsText);
        }

        private string GetDeclarationsText(VBAParser.ConstStmtContext consts)
        {
            var keyword = string.Empty;
            if (consts.visibility() != null)
            {
                keyword += consts.visibility().GetText() + ' ';
            }

            keyword += consts.CONST().GetText() + ' ';

            var newContent = new StringBuilder();
            foreach (var constant in consts.constSubStmt())
            {
                newContent.AppendLine(keyword + constant.GetText());
            }

            return newContent.ToString();
        }

        private string GetDeclarationsText(VBAParser.VariableStmtContext variables)
        {
            var keyword = string.Empty;
            if (variables.DIM() != null)
            {
                keyword += Tokens.Dim + ' ';
            }
            else if (variables.visibility() != null)
            {
                keyword += variables.visibility().GetText() + ' ';
            }
            else if (variables.STATIC() != null)
            {
                keyword += variables.STATIC().GetText() + ' ';
            }

            var newContent = new StringBuilder();
            foreach (var variable in variables.variableListStmt().variableSubStmt())
            {
                newContent.AppendLine(keyword + variable.GetText());
            }

            return newContent.ToString();
        }

        public override string Description(IInspectionResult result) => InspectionsUI.SplitMultipleDeclarationsQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}