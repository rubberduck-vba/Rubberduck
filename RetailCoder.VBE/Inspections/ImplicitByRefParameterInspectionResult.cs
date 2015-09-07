using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ImplicitByRefParameterInspectionResult : CodeInspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ImplicitByRefParameterInspectionResult(string inspection, CodeInspectionSeverity type, QualifiedContext<VBAParser.ArgContext> qualifiedContext)
            : base(inspection,type, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
            // array parameters & paramarrays must be passed by reference
            var context = (VBAParser.ArgContext) Context;
            if ((context.LPAREN() != null && context.RPAREN() != null) || context.PARAMARRAY() != null)
            {
                _quickFixes = new[]
                {
                    new ImplicitByRefParameterQuickFix(Context, QualifiedSelection, RubberduckUI.Inspections_PassParamByRefExplicitly, Tokens.ByRef), 
                };
            }
            else
            {
                _quickFixes = new[]
                {
                    new ImplicitByRefParameterQuickFix(Context, QualifiedSelection, RubberduckUI.Inspections_PassParamByRefExplicitly, Tokens.ByRef), 
                    new ImplicitByRefParameterQuickFix(Context, QualifiedSelection, RubberduckUI.Inspections_PassParamByValue, Tokens.ByVal), 
                };
            }
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
    }

    public class ImplicitByRefParameterQuickFix : CodeInspectionQuickFix
    {
        private readonly string _newToken;

        public ImplicitByRefParameterQuickFix(ParserRuleContext context, QualifiedSelection selection, string description, string newToken) 
            : base(context, selection, description)
        {
            _newToken = newToken;
        }

        public override void Fix()
        {
            var parameter = Context.GetText();
            var newContent = string.Concat(_newToken, " ", parameter);
            var selection = Selection.Selection;

            var module = Selection.QualifiedName.Component.CodeModule;
            var lines = module.get_Lines(selection.StartLine, selection.LineCount);

            var result = lines.Replace(parameter, newContent);
            module.ReplaceLine(selection.StartLine, result);
        }
    }
}