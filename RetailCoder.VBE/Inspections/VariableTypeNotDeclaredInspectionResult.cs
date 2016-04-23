using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class VariableTypeNotDeclaredInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public VariableTypeNotDeclaredInspectionResult(IInspection inspection, Declaration target)
            : base(inspection, target)
        {
            _quickFixes = new CodeInspectionQuickFix[]
            {
                new DeclareAsExplicitVariantQuickFix(Context, QualifiedSelection), 
                new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.ImplicitVariantDeclarationInspectionResultFormat, 
                    Target.DeclarationType,
                    Target.IdentifierName);
            }
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get {return _quickFixes; } }
    }

    public class DeclareAsExplicitVariantQuickFix : CodeInspectionQuickFix 
    {
        public DeclareAsExplicitVariantQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.DeclareAsExplicitVariantQuickFix)
        {
        }

        public override void Fix()
        {
            var codeModule = Selection.QualifiedName.Component.CodeModule;
            var codeLine = codeModule.Lines[Selection.Selection.StartLine, Selection.Selection.LineCount];

            // methods return empty string if soft-cast context is null - just concat results:
            string originalInstruction;
            
            var fix = DeclareExplicitVariant(Context as VBAParser.VariableSubStmtContext, out originalInstruction);

            if (string.IsNullOrEmpty(originalInstruction))
            {
                fix = DeclareExplicitVariant(Context.Parent as VBAParser.ConstSubStmtContext, out originalInstruction);
            }

            if (string.IsNullOrEmpty(originalInstruction))
            {
                fix = DeclareExplicitVariant(Context as VBAParser.ArgContext, out originalInstruction);
            }

            if (string.IsNullOrEmpty(originalInstruction))
            {
                return;
            }

            var fixedCodeLine = codeLine.Replace(originalInstruction, fix);
            codeModule.ReplaceLine(Selection.Selection.StartLine, fixedCodeLine);
        }

        private string DeclareExplicitVariant(VBAParser.VariableSubStmtContext context, out string instruction)
        {
            if (context == null)
            {
                instruction = null;
                return null;
            }

            instruction = context.GetText();
            return instruction + ' ' + Tokens.As + ' ' + Tokens.Variant;
        }

        private string DeclareExplicitVariant(VBAParser.ArgContext context, out string instruction)
        {
            if (context == null)
            {
                instruction = null;
                return null;
            }

            instruction = context.GetText();
            return instruction + ' ' + Tokens.As + ' ' + Tokens.Variant;
        }

        private string DeclareExplicitVariant(VBAParser.ConstSubStmtContext context, out string instruction)
        {
            if (context == null)
            {
                instruction = null;
                return null;
            }

            var parent = (VBAParser.ConstStmtContext)context.Parent;
            instruction = parent.GetText();

            var constant = context.GetText();
            var replacement = context.ambiguousIdentifier().GetText() + ' '
                         + Tokens.As + ' ' + Tokens.Variant + ' '
                         + context.EQ().GetText() + ' '
                         + context.valueStmt().GetText();

            var result = instruction.Replace(constant, replacement);
            return result;
        }
    }
}