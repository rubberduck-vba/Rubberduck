using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class RepaceEmptyStringLiteralStatementQuickFix : IQuickFix
    {
        public RepaceEmptyStringLiteralStatementQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.EmptyStringLiteralInspectionQuickFix)
        {
        }

        public void Fix(IInspectionResult result)
        {
            var module = Selection.QualifiedName.Component.CodeModule;
            if (module == null)
            {
                return;
            }

            var literal = (VBAParser.LiteralExpressionContext)Context;
            var newCodeLines = module.GetLines(literal.Start.Line, 1).Replace("\"\"", "vbNullString");

            module.ReplaceLine(literal.Start.Line, newCodeLines);
        }
    }
}