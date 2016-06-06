using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public sealed class ImplicitVariantReturnTypeInspectionResult : InspectionResultBase
    {
        private readonly string _identifierName;
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ImplicitVariantReturnTypeInspectionResult(IInspection inspection, string identifierName, QualifiedContext<ParserRuleContext> qualifiedContext, Declaration target)
            : base(inspection, qualifiedContext.ModuleName, qualifiedContext.Context, target)
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

        public override NavigateCodeEventArgs GetNavigationArgs()
        {
            return new NavigateCodeEventArgs(Target);
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
            var procedure = Context.GetText();
            var indexOfLastClosingParen = procedure.LastIndexOf(')');

            var result = indexOfLastClosingParen == procedure.Length
                ? procedure + ' ' + Tokens.As + ' ' + Tokens.Variant
                : procedure.Insert(procedure.LastIndexOf(')') + 1, ' ' + Tokens.As + ' ' + Tokens.Variant);
            
            var module = Selection.QualifiedName.Component.CodeModule;
            var selection = Context.GetSelection();

            module.DeleteLines(selection.StartLine, selection.LineCount);
            module.InsertLines(selection.StartLine, result);
        }

        private string GetSignature(VBAParser.FunctionStmtContext context)
        {
            if (context == null)
            {
                return null;
            }

            var @static = context.STATIC() == null ? string.Empty : context.STATIC().GetText() + ' ';
            var keyword = context.FUNCTION().GetText() + ' ';
            var args = context.argList() == null ? "()" : context.argList().GetText() + ' ';
            var asTypeClause = context.asTypeClause() == null ? string.Empty : context.asTypeClause().GetText();
            var visibility = context.visibility() == null ? string.Empty : context.visibility().GetText() + ' ';

            return visibility + @static + keyword + context.functionName().identifier().GetText() + args + asTypeClause;
        }

        private string GetSignature(VBAParser.PropertyGetStmtContext context)
        {
            if (context == null)
            {
                return null;
            }

            var @static = context.STATIC() == null ? string.Empty : context.STATIC().GetText() + ' ';
            var keyword = context.PROPERTY_GET().GetText() + ' ';
            var args = context.argList() == null ? "()" : context.argList().GetText() + ' ';
            var asTypeClause = context.asTypeClause() == null ? string.Empty : context.asTypeClause().GetText();
            var visibility = context.visibility() == null ? string.Empty : context.visibility().GetText() + ' ';

            return visibility + @static + keyword + context.functionName().identifier().GetText() + args + asTypeClause;
        }

        private string GetSignature(VBAParser.DeclareStmtContext context)
        {
            if (context == null)
            {
                return null;
            }

            var args = context.argList() == null ? "()" : context.argList().GetText() + ' ';
            var asTypeClause = context.asTypeClause() == null ? string.Empty : context.asTypeClause().GetText();

            return args + asTypeClause;
        }
    }
}
