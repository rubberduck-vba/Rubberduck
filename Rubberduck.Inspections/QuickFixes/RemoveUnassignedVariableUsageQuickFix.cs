using System;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    [Undocumented] // todo: use?
    public class RemoveUnassignedVariableUsageQuickFix : IQuickFix
    {
        public RemoveUnassignedVariableUsageQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.RemoveUnassignedVariableUsageQuickFix)
        {
        }

        public void Fix(IInspectionResult result)
        {
            var module = Selection.QualifiedName.Component.CodeModule;
            var selection = Selection.Selection;

            var originalCodeLines = module.GetLines(selection.StartLine, selection.LineCount)
                .Replace(Environment.NewLine, " ")
                .Replace("_", string.Empty);

            var originalInstruction = Context.GetText();
            module.DeleteLines(selection.StartLine, selection.LineCount);

            var newInstruction = InspectionsUI.Inspections_UnassignedVariableTodo;
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