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
        public ParameterCanBeByValInspectionResult(string inspection, CodeInspectionSeverity type,
            ParserRuleContext context, QualifiedMemberName qualifiedName)
            : base(inspection, type, qualifiedName.QualifiedModuleName, context)
        {
        }

        public override IDictionary<string, Action> GetQuickFixes()
        {
            return new Dictionary<string, Action>
            {
                {RubberduckUI.Inspections_PassParamByValue, PassParameterByValue}
            };
        }

        private void PassParameterByValue()
        {
            var parameter = Context.Parent.GetText();
            var newContent = string.Concat(Tokens.ByVal, " ", parameter.Replace(Tokens.ByRef, string.Empty).Trim());
            var selection = QualifiedSelection.Selection;

            var module = QualifiedName.Component.CodeModule;
            var lines = module.get_Lines(selection.StartLine, selection.LineCount);

            var result = lines.Replace(parameter, newContent);
            module.ReplaceLine(selection.StartLine, result);
        }
    }
}