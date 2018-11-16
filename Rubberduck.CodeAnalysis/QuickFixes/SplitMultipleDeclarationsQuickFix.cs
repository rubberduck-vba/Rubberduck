using System;
using System.Text;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class SplitMultipleDeclarationsQuickFix : QuickFixBase
    {
        public SplitMultipleDeclarationsQuickFix()
            : base(typeof(MultipleDeclarationsInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var context = result.Context is VBAParser.ConstStmtContext
                ? result.Context
                : (ParserRuleContext)result.Context.Parent;

            string declarationsText;
            switch (context)
            {
                case VBAParser.ConstStmtContext consts:
                    declarationsText = GetDeclarationsText(consts);
                    break;
                case VBAParser.VariableStmtContext variables:
                    declarationsText = GetDeclarationsText(variables);
                    break;
                default:
                    throw new NotSupportedException();
            }

            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);
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

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.SplitMultipleDeclarationsQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}