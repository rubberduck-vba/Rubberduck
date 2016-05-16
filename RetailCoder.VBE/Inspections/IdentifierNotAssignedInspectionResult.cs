using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class IdentifierNotAssignedInspectionResult : IdentifierNotUsedInspectionResult
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;
        private readonly Declaration _target;

        public IdentifierNotAssignedInspectionResult(IInspection inspection, Declaration target,
            ParserRuleContext context, QualifiedModuleName qualifiedName)
            : base(inspection, target, context, qualifiedName)
        {
            _target = target;
            _quickFixes = new CodeInspectionQuickFix[]
            {
                new RemoveUnassignedIdentifierQuickFix(Context, QualifiedSelection), 
                new IgnoreOnceQuickFix(context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override string Description
        {
            get { return string.Format(InspectionsUI.VariableNotAssignedInspectionResultFormat, _target.IdentifierName); }
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
    }

    public class RemoveUnassignedIdentifierQuickFix : CodeInspectionQuickFix
    {
        public RemoveUnassignedIdentifierQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.RemoveUnassignedIdentifierQuickFix)
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

            var newInstruction = string.Empty;
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