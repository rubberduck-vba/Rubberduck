using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class SpecifyExplicitPublicModifierQuickFix : CodeInspectionQuickFix
    {
        public SpecifyExplicitPublicModifierQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.SpecifyExplicitPublicModifierQuickFix)
        {
        }

        public override void Fix()
        {
            var selection = Context.GetSelection();
            var module = Selection.QualifiedName.Component.CodeModule;
            {
                var signatureLine = selection.StartLine;

                var oldContent = module.GetLines(signatureLine, 1);
                var newContent = Tokens.Public + ' ' + oldContent;

                module.DeleteLines(signatureLine);
                module.InsertLines(signatureLine, newContent);
            }
        }
    }
}