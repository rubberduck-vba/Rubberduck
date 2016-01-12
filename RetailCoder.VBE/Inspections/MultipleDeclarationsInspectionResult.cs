using System;
using System.Collections.Generic;
using System.Text;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class MultipleDeclarationsInspectionResult : CodeInspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public MultipleDeclarationsInspectionResult(IInspection inspection, string result, QualifiedContext<ParserRuleContext> qualifiedContext)
            : base(inspection, result, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
            _quickFixes = new CodeInspectionQuickFix[]
            {
                new SplitMultipleDeclarationsQuickFix(Context, QualifiedSelection), 
                new IgnoreOnceQuickFix(qualifiedContext.Context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get {return _quickFixes; } }

        private new QualifiedSelection QualifiedSelection
        {
            get
            {
                ParserRuleContext context;
                if (Context is VBAParser.ConstStmtContext)
                {
                    context = Context;
                }
                else
                {
                    context = Context.Parent as ParserRuleContext;
                }
                var selection = context.GetSelection();
                return new QualifiedSelection(QualifiedName, selection);
            }
        }
    }

    public class SplitMultipleDeclarationsQuickFix : CodeInspectionQuickFix
    {
        public SplitMultipleDeclarationsQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, RubberduckUI.Inspections_SplitDeclarations)
        {
        }

        public override void Fix()
        {
            var newContent = new StringBuilder();
            var selection = Selection.Selection;
            var keyword = string.Empty;

            var variables = Context.Parent as VBAParser.VariableStmtContext;
            if (variables != null)
            {
                if (variables.DIM() != null)
                {
                    keyword += Tokens.Dim + ' ';
                }
                else if (variables.visibility() != null)
                {
                    keyword += variables.visibility().GetText() + ' ';
                }
                else if (variables.STATIC() != null)
                {
                    keyword += variables.STATIC().GetText() + ' ';
                }

                foreach (var variable in variables.variableListStmt().variableSubStmt())
                {
                    newContent.AppendLine(keyword + variable.GetText());
                }
            }

            var consts = Context as VBAParser.ConstStmtContext;
            if (consts != null)
            {
                var keywords = string.Empty;

                if (consts.visibility() != null)
                {
                    keywords += consts.visibility().GetText() + ' ';
                }

                keywords += consts.CONST().GetText() + ' ';

                foreach (var constant in consts.constSubStmt())
                {
                    newContent.AppendLine(keywords + constant.GetText());
                }
            }

            var module = Selection.QualifiedName.Component.CodeModule;
            module.ReplaceLine(selection.StartLine, newContent.ToString());
        }
    }
}