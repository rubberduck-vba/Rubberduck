using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;

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
                if (Context is VBParser.ConstStmtContext)
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

            var variables = Context.Parent as VBParser.VariableStmtContext;
            if (variables != null)
            {
                if (variables.DIM() != null)
                {
                    keyword += Tokens.Dim + ' ';
                }
                else if(variables.Visibility() != null)
                {
                    keyword += variables.Visibility().GetText() + ' '; 
                }
                else if (variables.STATIC() != null)
                {
                    keyword += variables.STATIC().GetText() + ' ';
                }

                foreach (var variable in variables.VariableListStmt().VariableSubStmt())
                {
                    newContent.AppendLine(keyword + variable.GetText());
                }
            }

            var consts = Context as VBParser.ConstStmtContext;
            if (consts != null)
            {
                var keywords = string.Empty;

                if (consts.Visibility() != null)
                {
                    keywords += consts.Visibility().GetText() + ' ';
                }

                keywords += consts.CONST().GetText() + ' ';

                foreach (var constant in consts.ConstSubStmt())
                {
                    newContent.AppendLine(keywords + constant.GetText());
                }
            }

            var module = vbe.FindCodeModules(QualifiedName).First();
            module.ReplaceLine(selection.StartLine, newContent.ToString());
        }
    }
}