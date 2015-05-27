using System;
using System.Collections.Generic;
using System.Diagnostics;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBA.Nodes;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class ImplicitVariantReturnTypeInspectionResult : CodeInspectionResultBase
    {
        public ImplicitVariantReturnTypeInspectionResult(string name, CodeInspectionSeverity severity, 
            QualifiedContext<ParserRuleContext> qualifiedContext)
            : base(name, severity, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
        }

        public override IDictionary<string, Action> GetQuickFixes()
        {
            return
                new Dictionary<string, Action>
                {
                    {RubberduckUI.Inspections_ReturnExplicitVariant, ReturnExplicitVariant}
                };
        }

        private void ReturnExplicitVariant()
        {
            // note: turns a multiline signature into a one-liner signature.
            // bug: removes all comments.

            var node = GetNode(Context);
            var signature = node.Signature.TrimEnd();

            var procedure = Context.GetText();
            var result = procedure.Replace(signature, signature + ' ' + Tokens.As + ' ' + Tokens.Variant);
            
            var module = QualifiedName.Component.CodeModule;
            var selection = Context.GetSelection();

            module.DeleteLines(selection.StartLine, selection.LineCount);
            module.InsertLines(selection.StartLine, result);
        }

        private ProcedureNode GetNode(ParserRuleContext context)
        {
            var result = GetNode(context as VBAParser.FunctionStmtContext);
            if (result != null) return result;
            
            result = GetNode(context as VBAParser.PropertyGetStmtContext);
            Debug.Assert(result != null, "result != null");

            return result;
        }

        private ProcedureNode GetNode(VBAParser.FunctionStmtContext context)
        {
            if (context == null)
            {
                return null;
            }

            var scope = QualifiedName.ToString();
            var localScope = scope + "." + context.ambiguousIdentifier().GetText();
            return new ProcedureNode(context, scope, localScope);
        }

        private ProcedureNode GetNode(VBAParser.PropertyGetStmtContext context)
        {
            if (context == null)
            {
                return null;
            }

            var scope = QualifiedName.ToString();
            var localScope = scope + "." + context.ambiguousIdentifier().GetText();
            return new ProcedureNode(context, scope, localScope);
        }
    }
}