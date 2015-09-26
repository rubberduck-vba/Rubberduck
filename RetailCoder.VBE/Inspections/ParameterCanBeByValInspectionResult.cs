using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ParameterCanBeByValInspectionResult : CodeInspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ParameterCanBeByValInspectionResult(IInspection inspection, string result, ParserRuleContext context, QualifiedMemberName qualifiedName)
            : base(inspection, result, qualifiedName.QualifiedModuleName, context)
        {
            _quickFixes = new[]
            {
                new PassParameterByValueQuickFix(Context, QualifiedSelection), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
    }

    public class PassParameterByValueQuickFix : CodeInspectionQuickFix
    {
        public PassParameterByValueQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, RubberduckUI.Inspections_PassParamByValue)
        {
        }

        public override void Fix()
        {
            var parameter = Context.Parent.GetText();
            var newContent = string.Concat(Tokens.ByVal, " ", parameter.Replace(Tokens.ByRef, string.Empty).Trim());
            var selection = Selection.Selection;

            var module = Selection.QualifiedName.Component.CodeModule;
            var lines = module.get_Lines(selection.StartLine, selection.LineCount);

            var result = lines.Replace(parameter, newContent);
            module.ReplaceLine(selection.StartLine, result);
        }
    }
}