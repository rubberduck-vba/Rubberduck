using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.VBEditor;
using System;
using System.Linq;

namespace Rubberduck.Inspections.QuickFixes
{
    internal class RemoveOptionBaseStatementQuickFix : QuickFixBase
    {
        public RemoveOptionBaseStatementQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.RemoveOptionBaseStatementQuickFix)
        {
        }

        public override void Fix()
        {
            var module = Selection.QualifiedName.Component.CodeModule;
            var lines = module.GetLines(Selection.Selection).Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            var newContent = Selection.Selection.LineCount != 1
                ? lines[0].Remove(Selection.Selection.StartColumn - 1)
                : lines[0].Remove(Selection.Selection.StartColumn - 1, Selection.Selection.EndColumn - Selection.Selection.StartColumn);
            
            if (Selection.Selection.LineCount != 1)
            {
                newContent += lines.Last().Remove(0, Selection.Selection.EndColumn - 1);
            }

            module.DeleteLines(Selection.Selection);
            module.InsertLines(Selection.Selection.StartLine, newContent);
        }
    }
}