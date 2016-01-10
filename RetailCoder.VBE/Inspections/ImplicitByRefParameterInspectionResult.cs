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

        public ImplicitByRefParameterInspectionResult(IInspection inspection, string result, QualifiedContext<VBAParser.ArgContext> qualifiedContext)
            : base(inspection, result, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
            _quickFixes = new CodeInspectionQuickFix[]
                {
                    new ImplicitByRefParameterQuickFix(Context, QualifiedSelection, RubberduckUI.Inspections_PassParamByRefExplicitly, Tokens.ByRef), 
                    new IgnoreOnceQuickFix(qualifiedContext.Context, QualifiedSelection, Inspection.AnnotationName), 
                };
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