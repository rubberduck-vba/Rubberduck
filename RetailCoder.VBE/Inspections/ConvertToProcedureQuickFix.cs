using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Rubberduck.Inspections
{
    public class ConvertToProcedureQuickFix : CodeInspectionQuickFix
    {
        private readonly IEnumerable<string> _returnStatements;

        public ConvertToProcedureQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : this(context, selection, new List<string>())
        {
        }

        public ConvertToProcedureQuickFix(ParserRuleContext context, QualifiedSelection selection, IEnumerable<string> returnStatements)
            : base(context, selection, RubberduckUI.Inspections_ConvertFunctionToProcedure)
        {
            _returnStatements = returnStatements;
        }

        public override void Fix()
        {
            var context = (VBAParser.FunctionStmtContext)Context;
            var visibility = context.visibility() == null ? string.Empty : context.visibility().GetText() + ' ';
            var name = ' ' + context.ambiguousIdentifier().GetText();
            var args = context.argList().GetText();
            var asType = context.asTypeClause() == null ? string.Empty : ' ' + context.asTypeClause().GetText();

            var oldSignature = visibility + Tokens.Function + name + args + asType;
            var newSignature = visibility + Tokens.Sub + name + args;

            var procedure = Context.GetText();
            string noReturnStatements = procedure;
            _returnStatements.ToList().ForEach(returnStatement =>
                noReturnStatements = Regex.Replace(noReturnStatements, @"[ \t\f]*" + returnStatement + @"[ \t\f]*\r?\n?", ""));
            var result = noReturnStatements.Replace(oldSignature, newSignature)
                .Replace(Tokens.End + ' ' + Tokens.Function, Tokens.End + ' ' + Tokens.Sub)
                .Replace(Tokens.Exit + ' ' + Tokens.Function, Tokens.Exit + ' ' + Tokens.Sub);

            var module = Selection.QualifiedName.Component.CodeModule;
            var selection = Context.GetSelection();

            module.DeleteLines(selection.StartLine, selection.LineCount);
            module.InsertLines(selection.StartLine, result);
        }
    }
}
