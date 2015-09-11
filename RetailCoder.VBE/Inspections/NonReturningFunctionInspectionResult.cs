using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class NonReturningFunctionInspectionResult : CodeInspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public NonReturningFunctionInspectionResult(string inspection, CodeInspectionSeverity type, QualifiedContext<ParserRuleContext> qualifiedContext, bool isInterfaceImplementation)
            : base(inspection, type, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
            _quickFixes = isInterfaceImplementation 
                ? new CodeInspectionQuickFix[] { }
                : new[]
                {
                    new ConvertToProcedureQuickFix(Context, QualifiedSelection),
                };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
    }

    public class ConvertToProcedureQuickFix : CodeInspectionQuickFix
    {
        public ConvertToProcedureQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, RubberduckUI.Inspections_ConvertFunctionToProcedure)
        {
        }

        public override void Fix()
        {
            var context = (VBAParser.FunctionStmtContext) Context;
            var visibility = context.visibility() == null ? string.Empty : context.visibility().GetText() + ' ';
            var name = ' ' + context.ambiguousIdentifier().GetText();
            var args = context.argList().GetText();
            var asType = context.asTypeClause() == null ? string.Empty : ' ' + context.asTypeClause().GetText();

            var oldSignature = visibility + Tokens.Function + name + args + asType;
            var newSignature = visibility + Tokens.Sub + name + args;

            var procedure = Context.GetText();
            var result = procedure.Replace(oldSignature, newSignature)
                .Replace(Tokens.End + ' ' + Tokens.Function, Tokens.End + ' ' + Tokens.Sub)
                .Replace(Tokens.Exit + ' ' + Tokens.Function, Tokens.Exit + ' ' + Tokens.Sub);

            var module = Selection.QualifiedName.Component.CodeModule;
            var selection = Context.GetSelection();

            module.DeleteLines(selection.StartLine, selection.LineCount);
            module.InsertLines(selection.StartLine, result);
        }
    }
}