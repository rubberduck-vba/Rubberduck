using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class AssignedByValParameterInspectionResult : CodeInspectionResultBase
    {
        public AssignedByValParameterInspectionResult(string inspection, CodeInspectionSeverity type,
            ParserRuleContext context, QualifiedMemberName qualifiedName)
            : base(inspection, type, qualifiedName.QualifiedModuleName, context)
        {
        }

        public override IDictionary<string, Action> GetQuickFixes()
        {
            return new Dictionary<string, Action>
            {
                {"Pass parameter by reference", PassParameterByReference}
                //,{"Introduce local variable", IntroduceLocalVariable}
            };
        }

        //note: vbe parameter is not used
        private void PassParameterByReference()
        {
            var parameter = Context.GetText();
            var newContent = string.Concat(Tokens.ByRef, " ", parameter.Replace(Tokens.ByVal, string.Empty).Trim());
            var selection = QualifiedSelection.Selection;

            var module = QualifiedName.Component.CodeModule;
            var lines = module.get_Lines(selection.StartLine, selection.LineCount);

            var result = lines.Replace(parameter, newContent);
            module.ReplaceLine(selection.StartLine, result);
        }
    }
}