using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class ConvertToProcedureQuickFix : QuickFixBase
    {
        private readonly Declaration _target;

        public ConvertToProcedureQuickFix(ParserRuleContext context, QualifiedSelection selection, Declaration target)
            : base(context, selection, InspectionsUI.ConvertFunctionToProcedureQuickFix)
        {
            _target = target;
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

            var functionName = Context is VBAParser.FunctionStmtContext
                ? ((VBAParser.FunctionStmtContext) Context).functionName()
                : ((VBAParser.PropertyGetStmtContext) Context).functionName();

            var token = functionContext != null
                ? Tokens.Function
                : Tokens.Property + ' ' + Tokens.Get;
            var endToken = token == Tokens.Function
                ? token
                : Tokens.Property;

            string visibility = context.visibility() == null ? string.Empty : context.visibility().GetText() + ' ';
            var name = ' ' + Identifier.GetName(functionName.identifier());
            var hasTypeHint = Identifier.GetTypeHintValue(functionName.identifier()) != null;

            string args = context.argList().GetText();
            string asType = context.asTypeClause() == null ? string.Empty : ' ' + context.asTypeClause().GetText();

            var oldSignature = visibility + token + name + (hasTypeHint ? Identifier.GetTypeHintValue(functionName.identifier()) : string.Empty) + args + asType;
            var newSignature = visibility + Tokens.Sub + name + args;

            var procedure = Context.GetText();
            var noReturnStatements = procedure;

            GetReturnStatements(_target).ToList().ForEach(returnStatement =>
                noReturnStatements = Regex.Replace(noReturnStatements, @"[ \t\f]*" + returnStatement + @"[ \t\f]*\r?\n?", ""));
            var result = noReturnStatements.Replace(oldSignature, newSignature)
                .Replace(Tokens.End + ' ' + endToken, Tokens.End + ' ' + Tokens.Sub)
                .Replace(Tokens.Exit + ' ' + endToken, Tokens.Exit + ' ' + Tokens.Sub);

            var module = Selection.QualifiedName.Component.CodeModule;
            var selection = Context.GetSelection();

            module.DeleteLines(selection.StartLine, selection.LineCount);
            module.InsertLines(selection.StartLine, result);
        }

        private IEnumerable<string> GetReturnStatements(Declaration declaration)
        {
            return declaration.References
                .Where(usage => IsReturnStatement(declaration, usage))
                .Select(usage => usage.Context.Parent.GetText());
        }

        private bool IsReturnStatement(Declaration declaration, IdentifierReference assignment)
        {
            return assignment.ParentScoping.Equals(declaration) && assignment.Declaration.Equals(declaration);
        }
    }
}
