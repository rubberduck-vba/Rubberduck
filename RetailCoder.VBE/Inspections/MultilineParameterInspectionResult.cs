using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.UI;
using Rubberduck.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class MultilineParameterInspectionResult : CodeInspectionResultBase
    {
        public MultilineParameterInspectionResult(string inspection, CodeInspectionSeverity severity, ParserRuleContext context, QualifiedMemberName qualifiedName)
            : base(inspection, severity, qualifiedName.QualifiedModuleName, context)
        {

        }

        public override IDictionary<string, Action> GetQuickFixes()
        {
            return new Dictionary<string, Action>
            {
                {RubberduckUI.Inspections_MultilineParameter, WriteParamOnOneLine}
            };
        }

        private void WriteParamOnOneLine()
        {
            var module = QualifiedName.Component.CodeModule;
            var selection = QualifiedSelection.Selection;

            var lines = module.Lines[selection.StartLine, selection.EndLine - selection.StartLine + 1];

            var startLine = module.Lines[selection.StartLine, 1];
            var endLine = module.Lines[selection.EndLine, 1];

            var adjustedStartColumn = selection.StartColumn - 1;
            var adjustedEndColumn = lines.Length - (endLine.Length - (selection.EndColumn > endLine.Length ? endLine.Length : selection.EndColumn - 1));

            var parameter = lines.Substring(adjustedStartColumn,
                adjustedEndColumn - adjustedStartColumn)
                .Replace("_", "")
                .RemoveExtraSpaces();

            var start = startLine.Remove(adjustedStartColumn);
            var end = lines.Remove(0, adjustedEndColumn);

            module.ReplaceLine(selection.StartLine, start + parameter + end);
            module.DeleteLines(selection.StartLine + 1, selection.EndLine - selection.StartLine);
        }
    }
}