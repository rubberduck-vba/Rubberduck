using System;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Vbe.Interop;

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
            : base(context, selection, InspectionsUI.ConvertFunctionToProcedureQuickFix)
        {
            _returnStatements = returnStatements;
        }

        public override void Fix()
        {
            dynamic functionContext = Context as VBAParser.FunctionStmtContext;
            dynamic propertyGetContext = Context as VBAParser.PropertyGetStmtContext;

            var context = functionContext ?? propertyGetContext;
            if (context == null)
            {
                throw new InvalidOperationException(string.Format(InspectionsUI.InvalidContextTypeInspectionFix, Context.GetType(), GetType()));
            }

            string token = functionContext != null
                ? Tokens.Function
                : Tokens.Property + ' ' + Tokens.Get;
            string endToken = token == Tokens.Function
                ? token
                : Tokens.Property;

            string visibility = context.visibility() == null ? string.Empty : context.visibility().GetText() + ' ';
            string name = ' ' + context.identifier().GetText();
            bool hasTypeHint = context.typeHint() != null;

            string args = context.argList().GetText();
            string asType = context.asTypeClause() == null ? string.Empty : ' ' + context.asTypeClause().GetText();

            string oldSignature = visibility + token + name + (hasTypeHint ? context.typeHint().GetText() : string.Empty) + args + asType;
            string newSignature = visibility + Tokens.Sub + name + args;

            string procedure = Context.GetText();
            string noReturnStatements = procedure;
            _returnStatements.ToList().ForEach(returnStatement =>
                noReturnStatements = Regex.Replace(noReturnStatements, @"[ \t\f]*" + returnStatement + @"[ \t\f]*\r?\n?", ""));
            string result = noReturnStatements.Replace(oldSignature, newSignature)
                .Replace(Tokens.End + ' ' + endToken, Tokens.End + ' ' + Tokens.Sub)
                .Replace(Tokens.Exit + ' ' + endToken, Tokens.Exit + ' ' + Tokens.Sub);

            CodeModule module = Selection.QualifiedName.Component.CodeModule;
            Selection selection = Context.GetSelection();

            module.DeleteLines(selection.StartLine, selection.LineCount);
            module.InsertLines(selection.StartLine, result);
        }
    }
}
