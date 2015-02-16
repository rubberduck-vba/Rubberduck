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

        private new VisualBasic6Parser.FunctionStmtContext Context { get { return base.Context as VisualBasic6Parser.FunctionStmtContext; } }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return new Dictionary<string, Action<VBE>>
            {
                {"Convert function to procedure", ConvertFunctionToProcedure}
            };
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

            var module = vbe.FindCodeModules(QualifiedName).First();
            var selection = Context.GetSelection();

            module.DeleteLines(selection.StartLine, selection.LineCount);
            module.InsertLines(selection.StartLine, result);
        }
    }
}