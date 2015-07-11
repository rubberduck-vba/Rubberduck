using System;
using System.Collections.Generic;
using System.Text;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Inspections
{
    public class MultipleDeclarationsInspectionResult : CodeInspectionResultBase
    {
        private readonly IRubberduckFactory<IRubberduckCodePane> _factory;
        
        public MultipleDeclarationsInspectionResult(string inspection, CodeInspectionSeverity type, 
            QualifiedContext<ParserRuleContext> qualifiedContext, IRubberduckFactory<IRubberduckCodePane> factory)
            : base(inspection, type, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
            _factory = factory;
        }

        public override IDictionary<string, Action> GetQuickFixes()
        {
            return new Dictionary<string, Action>
            {
                {RubberduckUI.Inspections_SplitDeclarations, SplitDeclarations},
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
                return new QualifiedSelection(QualifiedName, selection, _factory);
            }
        }

        private void SplitDeclarations()
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

            var module = QualifiedName.Component.CodeModule;
            module.ReplaceLine(selection.StartLine, newContent.ToString());
        }
    }
}