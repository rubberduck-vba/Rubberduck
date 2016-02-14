using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class UnassignedVariableUsageInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public UnassignedVariableUsageInspectionResult(IInspection inspection, 
            ParserRuleContext context, QualifiedModuleName qualifiedName)
            : base(inspection, qualifiedName, context)
        {
            _quickFixes = new CodeInspectionQuickFix[]
            {
                new RemoveUnassignedVariableUsageQuickFix(Context, QualifiedSelection),
                new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.UnassignedVariableUsageInspectionResultFormat, Target.IdentifierName);
            }
        }
    }

    public class RemoveUnassignedVariableUsageQuickFix : CodeInspectionQuickFix
    {
        public RemoveUnassignedVariableUsageQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, RubberduckUI.Inspections_RemoveUsageBreaksCode)
        {
        }

        public override void Fix()
        {
            var module = Selection.QualifiedName.Component.CodeModule;
            var selection = Selection.Selection;

            var originalCodeLines = module.get_Lines(selection.StartLine, selection.LineCount)
                .Replace(Environment.NewLine, " ")
                .Replace("_", string.Empty);

            var originalInstruction = Context.GetText();
            module.DeleteLines(selection.StartLine, selection.LineCount);

            var newInstruction = RubberduckUI.Inspections_UnassignedVariableTodo;
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