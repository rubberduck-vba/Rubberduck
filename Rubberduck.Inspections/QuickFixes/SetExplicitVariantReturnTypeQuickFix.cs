using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class SetExplicitVariantReturnTypeQuickFix : QuickFixBase
    {
        public SetExplicitVariantReturnTypeQuickFix(ParserRuleContext context, QualifiedSelection selection, string description) 
            : base(context, selection, description)
        {
        }

        public override void Fix()
        {
            var procedure = Context.GetText();
            // todo: verify that this isn't a bug / test with a procedure that contains parentheses in the body.
            var indexOfLastClosingParen = procedure.LastIndexOf(')');

            var result = indexOfLastClosingParen == procedure.Length
                ? procedure + ' ' + Tokens.As + ' ' + Tokens.Variant
                : procedure.Insert(procedure.LastIndexOf(')') + 1, ' ' + Tokens.As + ' ' + Tokens.Variant);
            
            var module = Selection.QualifiedName.Component.CodeModule;
            var selection = Context.GetSelection();

            module.DeleteLines(selection.StartLine, selection.LineCount);
            module.InsertLines(selection.StartLine, result);
        }
    }
}