using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.VBEditor;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.QuickFixes
{
    public class RemoveExplicitByRefModifierQuickFix : QuickFixBase
    {
        public RemoveExplicitByRefModifierQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.ObsoleteByRefModifierQuickFix)
        {
        }

        public IEnumerable<Declaration> InterfaceImplementationDeclarations { get; set; }

        public override void Fix()
        {
            FixDeclaration(Selection, Context);

            if (InterfaceImplementationDeclarations != null)
            {
                foreach (var declaration in InterfaceImplementationDeclarations)
                {
                    FixDeclaration(declaration.QualifiedSelection, declaration.Context);
                }
            }
        }

        private static void FixDeclaration(QualifiedSelection qualifiedSelection, ParserRuleContext context)
        {
            var module = qualifiedSelection.QualifiedName.Component.CodeModule;
            var selection = context.GetSelection();
            var originalCodeLines = module.GetLines(selection.StartLine, selection.LineCount);
            var originalInstruction = context.GetText();

            var newInstruction = originalInstruction.Replace(Tokens.ByRef + ' ', "");

            module.DeleteLines(selection.StartLine, selection.LineCount);

            var newCodeLines = originalCodeLines.Replace(originalInstruction, newInstruction);

            module.InsertLines(selection.StartLine, newCodeLines);
        }
    }
}
