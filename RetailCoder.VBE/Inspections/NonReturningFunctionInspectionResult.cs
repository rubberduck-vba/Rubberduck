using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections
{
    public class NonReturningFunctionInspectionResult : CodeInspectionResultBase
    {
        private readonly bool _isInterfaceImplementation;

        public NonReturningFunctionInspectionResult(string inspection, CodeInspectionSeverity type, QualifiedContext<ParserRuleContext> qualifiedContext, bool isInterfaceImplementation)
            : base(inspection, type, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
            _isInterfaceImplementation = isInterfaceImplementation;
        }

        private new VBAParser.FunctionStmtContext Context { get { return base.Context as VBAParser.FunctionStmtContext; } }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            var result = new Dictionary<string, Action<VBE>>();
            if (!_isInterfaceImplementation) // changing procedure type would break interface implementation
            {
                result.Add("Convert function to procedure", ConvertFunctionToProcedure);
            }

            return result;
        }

        private void ConvertFunctionToProcedure(VBE vbe)
        {
            var visibility = Context.visibility() == null ? string.Empty : Context.visibility().GetText() + ' ';
            var name = ' ' + Context.ambiguousIdentifier().GetText();
            var args = Context.argList().GetText();
            var asType = Context.asTypeClause() == null ? string.Empty : ' ' + Context.asTypeClause().GetText();

            var oldSignature = visibility + Tokens.Function + name + args + asType;
            var newSignature = visibility +  Tokens.Sub + name + args;

            var procedure = Context.GetText();
            var result = procedure.Replace(oldSignature, newSignature)
                .Replace(Tokens.End + ' ' + Tokens.Function, Tokens.End + ' ' + Tokens.Sub)
                .Replace(Tokens.Exit + ' ' + Tokens.Function, Tokens.Exit + ' ' + Tokens.Sub);

            var module = QualifiedName.Component.CodeModule;
            var selection = Context.GetSelection();

            module.DeleteLines(selection.StartLine, selection.LineCount);
            module.InsertLines(selection.StartLine, result);
        }
    }
}