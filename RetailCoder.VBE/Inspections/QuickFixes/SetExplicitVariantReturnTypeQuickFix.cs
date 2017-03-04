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
            var indexOfLastClosingParen = procedure.LastIndexOf(')');

            var result = indexOfLastClosingParen == procedure.Length
                ? procedure + ' ' + Tokens.As + ' ' + Tokens.Variant
                : procedure.Insert(procedure.LastIndexOf(')') + 1, ' ' + Tokens.As + ' ' + Tokens.Variant);
            
            var module = Selection.QualifiedName.Component.CodeModule;
            var selection = Context.GetSelection();

            module.DeleteLines(selection.StartLine, selection.LineCount);
            module.InsertLines(selection.StartLine, result);
        }

        private string GetSignature(VBAParser.FunctionStmtContext context)
        {
            if (context == null)
            {
                return null;
            }

            var @static = context.STATIC() == null ? string.Empty : context.STATIC().GetText() + ' ';
            var keyword = context.FUNCTION().GetText() + ' ';
            var args = context.argList() == null ? "()" : context.argList().GetText() + ' ';
            var asTypeClause = context.asTypeClause() == null ? string.Empty : context.asTypeClause().GetText();
            var visibility = context.visibility() == null ? string.Empty : context.visibility().GetText() + ' ';

            return visibility + @static + keyword + context.functionName().identifier().GetText() + args + asTypeClause;
        }

        private string GetSignature(VBAParser.PropertyGetStmtContext context)
        {
            if (context == null)
            {
                return null;
            }

            var @static = context.STATIC() == null ? string.Empty : context.STATIC().GetText() + ' ';
            var keyword = context.PROPERTY_GET().GetText() + ' ';
            var args = context.argList() == null ? "()" : context.argList().GetText() + ' ';
            var asTypeClause = context.asTypeClause() == null ? string.Empty : context.asTypeClause().GetText();
            var visibility = context.visibility() == null ? string.Empty : context.visibility().GetText() + ' ';

            return visibility + @static + keyword + context.functionName().identifier().GetText() + args + asTypeClause;
        }

        private string GetSignature(VBAParser.DeclareStmtContext context)
        {
            if (context == null)
            {
                return null;
            }

            var args = context.argList() == null ? "()" : context.argList().GetText() + ' ';
            var asTypeClause = context.asTypeClause() == null ? string.Empty : context.asTypeClause().GetText();

            return args + asTypeClause;
        }
    }
}