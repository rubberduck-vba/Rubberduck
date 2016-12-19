using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class RemoveExplicitLetStatementQuickFix : QuickFixBase
    {
        public RemoveExplicitLetStatementQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.RemoveObsoleteStatementQuickFix)
        {
        }

        public override void Fix()
        {
            var module = Selection.QualifiedName.Component.CodeModule;
            {
                if (module.IsWrappingNullReference)
                {
                    return;
                }

                var selection = Context.GetSelection();
                var context = (VBAParser.LetStmtContext)Context;

                // remove line continuations to compare against context:
                var originalCodeLines = module.GetLines(selection.StartLine, selection.LineCount)
                    .Replace("\r\n", " ")
                    .Replace("_", string.Empty);
                var originalInstruction = Context.GetText();

                var identifier = context.lExpression().GetText();
                var value = context.expression().GetText();

                module.DeleteLines(selection.StartLine, selection.LineCount);

                var newInstruction = identifier + " = " + value;
                var newCodeLines = originalCodeLines.Replace(originalInstruction, newInstruction);

                module.InsertLines(selection.StartLine, newCodeLines);
            }
        }
    }
}