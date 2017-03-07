using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class RepaceEmptyStringLiteralStatementQuickFix : QuickFixBase
    {
        public RepaceEmptyStringLiteralStatementQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.EmptyStringLiteralInspectionQuickFix)
        {
        }

        public override void Fix()
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