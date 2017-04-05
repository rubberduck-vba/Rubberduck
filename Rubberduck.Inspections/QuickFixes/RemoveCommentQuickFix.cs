using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class RemoveCommentQuickFix : IQuickFix
    {
        public RemoveCommentQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.RemoveObsoleteStatementQuickFix)
        { }

        public void Fix(IInspectionResult result)
        {
            var module = Selection.QualifiedName.Component.CodeModule;

            if (module.IsWrappingNullReference)
            {
                return;                
            }

            var start = Context.Start.Line;
            var commentLine = module.GetLines(start, Selection.Selection.LineCount);
            var newLine = commentLine.Substring(0, Context.Start.Column).TrimEnd();

            module.DeleteLines(start, Selection.Selection.LineCount);
            if (newLine.TrimStart().Length > 0)
            {
                module.InsertLines(start, newLine);
            }
        }
    }
}