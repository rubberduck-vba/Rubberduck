using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;

namespace Rubberduck.Inspections
{
    public class VariableNotAssignedInspetionResult : CodeInspectionResultBase
    {
        public VariableNotAssignedInspetionResult(string inspection, CodeInspectionSeverity type,
            ParserRuleContext context, QualifiedModuleName qualifiedName)
            : base(inspection, type, qualifiedName, context)
        {
        }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return
                new Dictionary<string, Action<VBE>>
                {
                    {"Remove unassigned variable (and usages)", RemoveVariableDeclaration}
                };
        }

        private void RemoveVariableDeclaration(VBE vbe)
        {
            var module = vbe.FindCodeModules(QualifiedName).First();
            var selection = QualifiedSelection.Selection;

            var originalCodeLines = module.get_Lines(selection.StartLine, selection.LineCount)
                .Replace("\r\n", " ")
                .Replace("_", string.Empty);

            var originalInstruction = Context.GetText();
            module.DeleteLines(selection.StartLine, selection.LineCount);

            var newInstruction = string.Empty;
            var newCodeLines = string.IsNullOrEmpty(newInstruction)
                ? string.Empty
                : originalCodeLines.Replace(originalInstruction, newInstruction);

            if (!string.IsNullOrEmpty(newCodeLines))
            {
                module.InsertLines(selection.StartLine, newCodeLines);
            }
        }

        private void RemoveVariableUsages(VBE vbe)
        {
            
        }
    }
}