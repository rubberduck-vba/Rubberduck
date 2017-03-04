using System.Text;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class SplitMultipleDeclarationsQuickFix : QuickFixBase
    {
        public SplitMultipleDeclarationsQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.SplitMultipleDeclarationsQuickFix)
        {
        }

        public override void Fix()
        {
            var newContent = new StringBuilder();
            var selection = Selection.Selection;
            var keyword = string.Empty;

            var variables = Context.Parent as VBAParser.VariableStmtContext;
            if (variables != null)
            {
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

                foreach (var variable in variables.variableListStmt().variableSubStmt())
                {
                    newContent.AppendLine(keyword + variable.GetText());
                }
            }

            var consts = Context as VBAParser.ConstStmtContext;
            if (consts != null)
            {
                var keywords = string.Empty;

                if (consts.visibility() != null)
                {
                    keywords += consts.visibility().GetText() + ' ';
                }

                keywords += consts.CONST().GetText() + ' ';

                foreach (var constant in consts.constSubStmt())
                {
                    newContent.AppendLine(keywords + constant.GetText());
                }
            }

            var module = Selection.QualifiedName.Component.CodeModule;
            module.DeleteLines(selection);
            module.InsertLines(selection.StartLine, newContent.ToString());
        }
    }
}