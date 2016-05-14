using Antlr4.Runtime;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class MakeSingleLineParameterQuickFix : CodeInspectionQuickFix
    {
        public MakeSingleLineParameterQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.MakeSingleLineParameterQuickFix)
        {
        }

        public override void Fix()
        {
            var module = Selection.QualifiedName.Component.CodeModule;
            var selection = Selection.Selection;

            var lines = module.Lines[selection.StartLine, selection.EndLine - selection.StartLine + 1];

            var startLine = module.Lines[selection.StartLine, 1];
            var endLine = module.Lines[selection.EndLine, 1];

            var adjustedStartColumn = selection.StartColumn - 1;
            var adjustedEndColumn = lines.Length - (endLine.Length - (selection.EndColumn > endLine.Length ? endLine.Length : selection.EndColumn - 1));

            var parameter = lines.Substring(adjustedStartColumn,
                adjustedEndColumn - adjustedStartColumn)
                .Replace("_", "")
                .RemoveExtraSpacesLeavingIndentation();

            var start = startLine.Remove(adjustedStartColumn);
            var end = lines.Remove(0, adjustedEndColumn);

            module.ReplaceLine(selection.StartLine, start + parameter + end);
            module.DeleteLines(selection.StartLine + 1, selection.EndLine - selection.StartLine);
        }
    }
}
