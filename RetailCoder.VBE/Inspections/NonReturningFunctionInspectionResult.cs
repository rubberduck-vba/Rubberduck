using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.Inspections
{
    public class NonReturningFunctionInspectionResult : CodeInspectionResultBase
    {
        public NonReturningFunctionInspectionResult(string inspection, CodeInspectionSeverity type, QualifiedContext<ParserRuleContext> qualifiedContext)
            : base(inspection, type, qualifiedContext.QualifiedName, qualifiedContext.Context)
        {
        }

        private new VBParser.FunctionStmtContext Context { get { return base.Context as VBParser.FunctionStmtContext; } }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return new Dictionary<string, Action<VBE>>
            {
                {"Convert function to procedure", ConvertFunctionToProcedure}
            };
        }

        private void ConvertFunctionToProcedure(VBE vbe)
        {
            var visibility = Context.Visibility() == null ? string.Empty : Context.Visibility().GetText() + ' ';
            var name = ' ' + Context.AmbiguousIdentifier().GetText();
            var args = Context.ArgList().GetText();
            var asType = Context.AsTypeClause() == null ? string.Empty : ' ' + Context.AsTypeClause().GetText();

            var oldSignature = visibility + Tokens.Function + name + args + asType;
            var newSignature = visibility +  Tokens.Sub + name + args;

            var procedure = Context.GetText();
            var result = procedure.Replace(oldSignature, newSignature)
                .Replace(Tokens.End + ' ' + Tokens.Function, Tokens.End + ' ' + Tokens.Sub)
                .Replace(Tokens.Exit + ' ' + Tokens.Function, Tokens.Exit + ' ' + Tokens.Sub);

            var module = vbe.FindCodeModules(QualifiedName).First();
            var selection = Context.GetSelection();

            module.DeleteLines(selection.StartLine, selection.LineCount);
            module.InsertLines(selection.StartLine, result);
        }
    }
}