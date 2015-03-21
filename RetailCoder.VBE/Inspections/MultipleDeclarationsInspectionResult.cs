using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections
{
    public class MultipleDeclarationsInspectionResult : CodeInspectionResultBase
    {
        public MultipleDeclarationsInspectionResult(string inspection, CodeInspectionSeverity type, 
            QualifiedContext<ParserRuleContext> qualifiedContext)
            : base(inspection, type, qualifiedContext.QualifiedName, qualifiedContext.Context)
        {
        }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return new Dictionary<string, Action<VBE>>
            {
                {"Separate multiple declarations into multiple instructions", SplitDeclarations},
            };
        }

        public override QualifiedSelection QualifiedSelection
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

        private void SplitDeclarations(VBE vbe)
        {
            var newContent = new StringBuilder();
            var selection = QualifiedSelection.Selection;
            string keyword = string.Empty;

            var variables = Context.Parent as VBAParser.VariableStmtContext;
            if (variables != null)
            {
                if (variables.DIM() != null)
                {
                    keyword += Tokens.Dim + ' ';
                }
                else if(variables.visibility() != null)
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

            var module = vbe.FindCodeModules(QualifiedName).First();
            module.ReplaceLine(selection.StartLine, newContent.ToString());
        }
    }
}