using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class UnassignedVariableUsageInspectionResult : CodeInspectionResultBase
    {
        public UnassignedVariableUsageInspectionResult(string inspection, CodeInspectionSeverity type,
            ParserRuleContext context, QualifiedModuleName qualifiedName)
            : base(inspection, type, qualifiedName, context)
        {
        }

        public override IDictionary<string, Action> GetQuickFixes()
        {
            return
                new Dictionary<string, Action>
                {
                    {"Remove usage (breaks code)", RemoveUsage}
                };
        }

        private void RemoveUsage()
        {
            var module = QualifiedName.Component.CodeModule;
            var selection = QualifiedSelection.Selection;

            var originalCodeLines = module.get_Lines(selection.StartLine, selection.LineCount)
                .Replace(Environment.NewLine, " ")
                .Replace("_", string.Empty);

            var originalInstruction = Context.GetText();
            module.DeleteLines(selection.StartLine, selection.LineCount);

            var newInstruction = "TODO";
            var newCodeLines = string.IsNullOrEmpty(newInstruction)
                ? string.Empty
                : originalCodeLines.Replace(originalInstruction, newInstruction);

            if (!string.IsNullOrEmpty(newCodeLines))
            {
                module.InsertLines(selection.StartLine, newCodeLines);
            }
        }
    }
}