using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing;
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
            var module = Selection.QualifiedName.Component.CodeModule;
            {
                string originalInstruction;

                // DeclareExplicitVariant() overloads return empty string if context is null
                Selection selection;
                var fix = DeclareExplicitVariant(Context as VBAParser.VariableSubStmtContext, out originalInstruction, out selection);

                if (string.IsNullOrEmpty(originalInstruction))
                {
                    fix = DeclareExplicitVariant(Context as VBAParser.ConstSubStmtContext, out originalInstruction, out selection);
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
                                    fix += part.GetText() + ' ' + Tokens.As + ' ' + Tokens.Variant + ' ';
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

        private string DeclareExplicitVariant(VBAParser.VariableSubStmtContext context, out string instruction, out Selection selection)
        {
            if (context == null)
            {
                instruction = null;
                selection = VBEditor.Selection.Empty;
                return null;
            }

            var parent = (ParserRuleContext)context.Parent.Parent;
            instruction = parent.GetText();
            selection = parent.GetSelection();

            var variable = context.GetText();
            var replacement = context.identifier().GetText() + ' ' + Tokens.As + ' ' + Tokens.Variant + ' ';

            var result = instruction.Replace(variable, replacement);
            return result;
        }

        private string DeclareExplicitVariant(VBAParser.ConstSubStmtContext context, out string instruction, out Selection selection)
        {
            if (context == null)
            {
                instruction = null;
                selection = VBEditor.Selection.Empty;
                return null;
            }

            var parent = (ParserRuleContext)context.Parent;
            instruction = parent.GetText();
            selection = parent.GetSelection(); 

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