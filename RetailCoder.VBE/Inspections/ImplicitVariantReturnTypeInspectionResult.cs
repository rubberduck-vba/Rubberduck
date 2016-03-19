using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.Nodes;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public sealed class ImplicitVariantReturnTypeInspectionResult : InspectionResultBase
    {
        private readonly string _identifierName;
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ImplicitVariantReturnTypeInspectionResult(IInspection inspection, string identifierName, QualifiedContext<ParserRuleContext> qualifiedContext)
            : base(inspection, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
            _identifierName = identifierName;
            _quickFixes = new CodeInspectionQuickFix[]
            {
                new SetExplicitVariantReturnTypeQuickFix(Context, QualifiedSelection, InspectionsUI.SetExplicitVariantReturnTypeQuickFix), 
                new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get {return _quickFixes; } }

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.ImplicitVariantReturnTypeInspectionResultFormat,
                    _identifierName);
            }
        }
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