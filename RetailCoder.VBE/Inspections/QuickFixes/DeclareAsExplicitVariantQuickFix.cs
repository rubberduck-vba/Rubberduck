using System;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class DeclareAsExplicitVariantQuickFix : QuickFixBase 
    {
        public DeclareAsExplicitVariantQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.DeclareAsExplicitVariantQuickFix)
        {
        }

        public override void Fix()
        {
            var module = Selection.QualifiedName.Component.CodeModule;
            var contextLines = module.GetLines(Context.GetSelection());
            var originalIndent = contextLines.Substring(0, contextLines.TakeWhile(c => c == ' ').Count());

            string originalInstruction;

            // DeclareExplicitVariant() overloads return empty string if context is null
            Selection selection;
            var fix = DeclareExplicitVariant(Context as VBAParser.VariableSubStmtContext, contextLines, out originalInstruction, out selection);
            if (!string.IsNullOrEmpty(fix))
            {
                // maintain original indentation for a variable declaration
                fix = originalIndent + fix;
            }

            if (string.IsNullOrEmpty(originalInstruction))
            {
                fix = DeclareExplicitVariant(Context as VBAParser.ConstSubStmtContext, contextLines, out originalInstruction, out selection);
            }

            if (string.IsNullOrEmpty(originalInstruction))
            {
                fix = DeclareExplicitVariant(Context as VBAParser.ArgContext, out originalInstruction, out selection);
            }

            if (string.IsNullOrEmpty(originalInstruction))
            {
                return;
            }

            module.DeleteLines(selection.StartLine, selection.LineCount);
            module.InsertLines(selection.StartLine, fix);
        }

        private string DeclareExplicitVariant(VBAParser.ArgContext context, out string instruction, out Selection selection)
        {
            if (context == null)
            {
                instruction = null;
                selection = VBEditor.Selection.Empty;
                return null;
            }

            var memberContext = (ParserRuleContext) context.Parent.Parent;
            selection = memberContext.GetSelection();
            instruction = memberContext.GetText();

            var fix = string.Empty;
            foreach (var child in memberContext.children)
            {
                if (child is VBAParser.ArgListContext)
                {
                    foreach (var tree in ((VBAParser.ArgListContext) child).children)
                    {
                        if (tree.Equals(context))
                        {
                            foreach (var part in context.children)
                            {
                                if (part is VBAParser.UnrestrictedIdentifierContext)
                                {
                                    fix += part.GetText() + ' ' + Tokens.As + ' ' + Tokens.Variant;
                                }
                                else
                                {
                                    fix += part.GetText();
                                }
                            }
                        }
                        else
                        {
                            fix += tree.GetText();
                        }
                    }
                }
                else
                {
                    fix += child.GetText();
                }
            }

            return fix;
        }

        private string DeclareExplicitVariant(VBAParser.VariableSubStmtContext context, string contextLines, out string instruction, out Selection selection)
        {
            if (context == null)
            {
                instruction = null;
                selection = VBEditor.Selection.Empty;
                return null;
            }

            var parent = (ParserRuleContext)context.Parent.Parent;
            selection = parent.GetSelection();
            instruction = contextLines.Substring(selection.StartColumn - 1);

            var variable = context.GetText();
            var replacement = context.identifier().GetText() + ' ' + Tokens.As + ' ' + Tokens.Variant;

            var insertIndex = instruction.IndexOf(variable, StringComparison.Ordinal);
            var result = instruction.Substring(0, insertIndex)
                         + replacement + instruction.Substring(insertIndex + variable.Length);
            return result;
        }

        private string DeclareExplicitVariant(VBAParser.ConstSubStmtContext context, string contextLines, out string instruction, out Selection selection)
        {
            if (context == null)
            {
                instruction = null;
                selection = VBEditor.Selection.Empty;
                return null;
            }

            var parent = (ParserRuleContext)context.Parent;
            selection = parent.GetSelection();
            instruction = contextLines.Substring(selection.StartColumn - 1);

            var constant = context.GetText();
            var replacement = context.identifier().GetText() + ' '
                              + Tokens.As + ' ' + Tokens.Variant + ' '
                              + context.EQ().GetText() + ' '
                              + context.expression().GetText();

            var insertIndex = instruction.IndexOf(constant, StringComparison.Ordinal);
            var result = instruction.Substring(0, insertIndex)
                         + replacement + instruction.Substring(insertIndex + constant.Length);
            return result;
        }
    }
}