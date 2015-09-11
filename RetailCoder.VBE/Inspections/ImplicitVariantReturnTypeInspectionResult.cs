using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.UI;
using Rubberduck.VBA.Nodes;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ImplicitVariantReturnTypeInspectionResult : CodeInspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ImplicitVariantReturnTypeInspectionResult(string name, CodeInspectionSeverity severity, QualifiedContext<ParserRuleContext> qualifiedContext)
            : base(name, severity, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
            _quickFixes = new[]
            {
                new SetExplicitVariantReturnTypeQuickFix(Context, QualifiedSelection, RubberduckUI.Inspections_ReturnExplicitVariant), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get {return _quickFixes; } }
    }

    public class SetExplicitVariantReturnTypeQuickFix : CodeInspectionQuickFix
    {
        public SetExplicitVariantReturnTypeQuickFix(ParserRuleContext context, QualifiedSelection selection, string description) 
            : base(context, selection, description)
        {
        }

        public override void Fix()
        {
            // note: turns a multiline signature into a one-liner signature.
            // bug: removes all comments.

            var node = GetNode(Context as VBAParser.FunctionStmtContext)
                    ?? GetNode(Context as VBAParser.PropertyGetStmtContext);

            var signature = node.Signature.TrimEnd();

            var procedure = Context.GetText();
            var result = procedure.Replace(signature, signature + ' ' + Tokens.As + ' ' + Tokens.Variant);
            
            var module = Selection.QualifiedName.Component.CodeModule;
            var selection = Context.GetSelection();

            module.DeleteLines(selection.StartLine, selection.LineCount);
            module.InsertLines(selection.StartLine, result);
        }

        private ProcedureNode GetNode(VBAParser.FunctionStmtContext context)
        {
            if (context == null)
            {
                return null;
            }

            var scope = Selection.QualifiedName.ToString();
            var localScope = scope + "." + context.ambiguousIdentifier().GetText();
            return new ProcedureNode(context, scope, localScope);
        }

        private ProcedureNode GetNode(VBAParser.PropertyGetStmtContext context)
        {
            if (context == null)
            {
                return null;
            }

            var scope = Selection.QualifiedName.ToString();
            var localScope = scope + "." + context.ambiguousIdentifier().GetText();
            return new ProcedureNode(context, scope, localScope);
        }
    }
}