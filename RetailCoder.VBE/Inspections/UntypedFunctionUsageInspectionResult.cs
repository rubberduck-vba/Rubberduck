using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class UntypedFunctionUsageInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public UntypedFunctionUsageInspectionResult(IInspection inspection, string result, QualifiedModuleName qualifiedName, ParserRuleContext context) 
            : base(inspection, result, qualifiedName, context)
        {
            _quickFixes = new CodeInspectionQuickFix[]
            {
                new UntypedFunctionUsageQuickFix(Context, QualifiedSelection), 
                new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
    }

    public class UntypedFunctionUsageQuickFix : CodeInspectionQuickFix
    {
        public UntypedFunctionUsageQuickFix(ParserRuleContext context, QualifiedSelection selection) 
            : base(context, selection, string.Format(RubberduckUI.QuickFixUseTypedFunction_, context.GetText()))
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
        }
    }
}