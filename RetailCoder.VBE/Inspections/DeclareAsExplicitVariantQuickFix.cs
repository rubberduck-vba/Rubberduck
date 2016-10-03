using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class DeclareAsExplicitVariantQuickFix : CodeInspectionQuickFix 
    {
        public DeclareAsExplicitVariantQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.DeclareAsExplicitVariantQuickFix)
        {
        }

        public override void Fix()
        {
            using (var module = Selection.QualifiedName.Component.CodeModule)
            {
                var codeLine = module.GetLines(Selection.Selection.StartLine, Selection.Selection.LineCount);

                // methods return empty string if soft-cast context is null - just concat results:
                string originalInstruction;

                var fix = DeclareExplicitVariant(Context as VBAParser.VariableSubStmtContext, out originalInstruction);

                if (string.IsNullOrEmpty(originalInstruction))
                {
                    fix = DeclareExplicitVariant(Context as VBAParser.ConstSubStmtContext, out originalInstruction);
                }

                if (string.IsNullOrEmpty(originalInstruction))
                {
                    fix = DeclareExplicitVariant(Context as VBAParser.ArgContext, out originalInstruction);
                }

                if (string.IsNullOrEmpty(originalInstruction))
                {
                    return;
                }

                var fixedCodeLine = codeLine.Remove(Context.Start.Column, originalInstruction.Length).Insert(Context.Start.Column, fix);
                module.ReplaceLine(Selection.Selection.StartLine, fixedCodeLine);
            }
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
            if (!context.children.Select(s => s.GetType()).Contains(typeof(VBAParser.ArgDefaultValueContext)))
            {
                return instruction + ' ' + Tokens.As + ' ' + Tokens.Variant;
            }

            var fix = string.Empty;
            var hasArgDefaultValue = false;
            foreach (var child in context.children)
            {
                if (child.GetType() == typeof(VBAParser.ArgDefaultValueContext))
                {
                    fix += Tokens.As + ' ' + Tokens.Variant + ' ';
                    hasArgDefaultValue = true;
                }

                fix += child.GetText();
            }

            return hasArgDefaultValue ? fix : fix + ' ' + Tokens.As + ' ' + Tokens.Variant;
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
            var replacement = context.identifier().GetText() + ' '
                              + Tokens.As + ' ' + Tokens.Variant + ' '
                              + context.EQ().GetText() + ' '
                              + context.expression().GetText();

            var result = instruction.Replace(constant, replacement);
            return result;
        }
    }
}