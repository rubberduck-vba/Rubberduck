using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class UntypedFunctionUsageInspectionResult : InspectionResultBase
    {
        private readonly IdentifierReference _reference;
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public UntypedFunctionUsageInspectionResult(IInspection inspection, IdentifierReference reference) 
            : base(inspection, reference.QualifiedModuleName, reference.Context)
        {
            _reference = reference;
            _quickFixes = new CodeInspectionQuickFix[]
            {
                new UntypedFunctionUsageQuickFix(Context, QualifiedSelection), 
                new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return string.Format(Inspection.Description, _reference.Declaration.IdentifierName); }
        }
    }

    public class UntypedFunctionUsageQuickFix : CodeInspectionQuickFix
    {
        public UntypedFunctionUsageQuickFix(ParserRuleContext context, QualifiedSelection selection) 
            : base(context, selection, string.Format(InspectionsUI.QuickFixUseTypedFunction_, context.GetText(), context.GetText() + "$"))
        {
        }

        public override void Fix()
        {
            var originalInstruction = Context.GetText();
            var newInstruction = originalInstruction + "$";
            var selection = Selection.Selection;

            var module = Selection.QualifiedName.Component.CodeModule;
            var lines = module.get_Lines(selection.StartLine, selection.LineCount);

            var result = lines.Replace(originalInstruction, newInstruction);
            module.ReplaceLine(selection.StartLine, result);
            // FIXME trigger reparse
        }
    }
}